VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmҽ���ӿڹ��� 
   Caption         =   "ҽ���ӿڹ���"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8715
   Icon            =   "frmҽ���ӿڹ���.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   8715
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.ImageList imgProp 
      Left            =   3690
      Top             =   3780
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
            Picture         =   "frmҽ���ӿڹ���.frx":1CFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgModul 
      Left            =   3690
      Top             =   1170
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
            Picture         =   "frmҽ���ӿڹ���.frx":4D7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgInterface 
      Left            =   30
      Top             =   750
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
            Picture         =   "frmҽ���ӿڹ���.frx":5BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ӿڹ���.frx":6E50
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwInterface 
      Height          =   4635
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgInterface"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ҽ����������"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrBlack 
      Left            =   3660
      Top             =   0
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
            Picture         =   "frmҽ���ӿڹ���.frx":80D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ӿڹ���.frx":82EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ӿڹ���.frx":8506
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ӿڹ���.frx":8958
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ӿڹ���.frx":8B72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrColor 
      Left            =   3090
      Top             =   0
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
            Picture         =   "frmҽ���ӿڹ���.frx":8D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ӿڹ���.frx":90DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ӿڹ���.frx":9430
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ӿڹ���.frx":9882
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ���ӿڹ���.frx":9A9C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   1244
      BandCount       =   1
      FixedOrder      =   -1  'True
      _CBWidth        =   8715
      _CBHeight       =   705
      _Version        =   "6.7.8988"
      Child1          =   "tbrTool"
      MinHeight1      =   645
      Width1          =   1575
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrTool 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgTbrBlack"
         HotImageList    =   "imgTbrColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��װ"
               Key             =   "Install"
               Object.ToolTipText     =   "��װҽ���ӿڲ���"
               Object.Tag             =   "��װ"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ж��"
               Key             =   "Uninstall"
               Object.ToolTipText     =   "ж��ҽ���ӿڲ���"
               Object.Tag             =   "ж��"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split0"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Start"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split1"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   5340
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   635
      SimpleText      =   $"frmҽ���ӿڹ���.frx":9CB6
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmҽ���ӿڹ���.frx":9CFD
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin MSComctlLib.ListView lvw����ģ�� 
      Height          =   2385
      Left            =   3630
      TabIndex        =   6
      Top             =   930
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4207
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgModul"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ģ��"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView lvw������ 
      Height          =   1815
      Left            =   3630
      TabIndex        =   4
      Top             =   3540
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgProp"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "˵��"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "�ӿڵ���"
      ForeColor       =   &H8000000E&
      Height          =   180
      Index           =   0
      Left            =   3630
      TabIndex        =   5
      Top             =   750
      Width           =   5040
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "������"
      ForeColor       =   &H8000000E&
      Height          =   180
      Index           =   1
      Left            =   3630
      MousePointer    =   7  'Size N S
      TabIndex        =   3
      Top             =   3360
      Width           =   5040
   End
   Begin VB.Image imgSplit 
      Height          =   4605
      Left            =   3540
      MousePointer    =   9  'Size W E
      Top             =   750
      Width           =   75
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuInterface 
      Caption         =   "�ӿ�(&I)"
      Begin VB.Menu mnuInterfaceInstall 
         Caption         =   "��װ(&I)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuInterfaceUninstall 
         Caption         =   "ж��(&U)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuInterfaceSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInterfaceStart 
         Caption         =   "����(&S)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&M)..."
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmҽ���ӿڹ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
Private mstrInsureUser As String
Private mstrInsureTablespace As String
Private mstrInsureName As String
Private mstrDemo As String
Private mstrComponent As String
Private mstrPath As String
Private mstrSQL As String
Private mstrUser As String              '�û���
Private mstrServer As String            '������
Private mblnDBA As Boolean              'DBA�û���������

Private mobjTest() As Object
Private mstrTest() As String
Private mobjConfigure As Object

Private mblnMove As Boolean
Private Type ����
    x As Double
    y As Double
End Type
Private Type_Scale As ����
'ֻ�������߻�DBA����Ȩ����ҽ���ӿڹ���İ�װ��ж��
'��ͨ�û�ֻ������нӿڵĵ���

Private Sub Form_Load()
    mstrUser = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "USER", "")
    mstrServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    
    mblnDBA = IsDBA()
    mnuInterfaceInstall.Visible = mblnDBA
    mnuInterfaceUninstall.Visible = mblnDBA
    mnuInterfaceSplit1.Visible = mblnDBA
    tbrTool.Buttons("Install").Visible = mblnDBA
    tbrTool.Buttons("Uninstall").Visible = mblnDBA
    tbrTool.Buttons("split0").Visible = mblnDBA
    
    Call LoadInterface
End Sub

Private Sub LoadInterface()
    Dim lvwItem As ListItem
    Dim rsInsure As New ADODB.Recordset
    On Error GoTo ErrHand
    
    'װ����ע��ҽ���ӿڵ�����
    mstrSQL = " Select A.���,A.����,B.���� As ҽ������,Nvl(B.����,0) ����" & _
              " From ������� A,zlInsureComponents B" & _
              " Where A.���=B.����" & _
              " Order By A.���"
    Call zlDatabase.OpenRecordset(rsInsure, mstrSQL, "װ����ע���ҽ���ӿ�")
    
    With rsInsure
        lvwInterface.ListItems.Clear
        Do While Not .EOF
            Set lvwItem = lvwInterface.ListItems.Add(, "K_" & !���, !���, , !���� + 1)
            lvwItem.SubItems(1) = Nvl(!����)
            lvwItem.SubItems(2) = Nvl(!ҽ������)
            lvwItem.Tag = !���� + 1
            .MoveNext
        Loop
    End With
    
    '����У�����õ���¼�����ʾ��ϸ��Ϣ��������ؿؼ�����ť����Ϊ������ѡ��״̬
    If Me.lvwInterface.ListItems.Count <> 0 Then
        Call lvwInterface_ItemClick(lvwInterface.ListItems(1))
    Else
        Call SetEnabled(False)
    End If
    Me.stbThis.Panels(2).Text = ""
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With Me.lvwInterface
        .Width = imgSplit.Left
        .Height = Me.ScaleHeight - stbThis.Height - .Top
    End With
    With lblNote(0)
        .Top = lvwInterface.Top
        .Left = imgSplit.Left + imgSplit.Width
        .Width = Me.ScaleWidth - .Left
    End With
    With Me.lvw����ģ��
        .Top = lblNote(0).Top + lblNote(0).Height
        .Left = lblNote(0).Left
        .Width = lblNote(0).Width
        .Height = lblNote(1).Top - .Top
    End With
    With lblNote(1)
        .Left = lblNote(0).Left
        .Width = lblNote(0).Width
    End With
    With Me.lvw������
        .Top = lblNote(1).Top + lblNote(1).Height
        .Left = lblNote(0).Left
        .Width = lblNote(0).Width
        .Height = Me.ScaleHeight - stbThis.Height - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intIndex As Integer, intCount As Integer
    On Error Resume Next
    
    '�ر����ж���
    intCount = UBound(mobjTest)
    If Err <> 0 Then intCount = -1
    
    For intIndex = 0 To intCount
        Call mobjTest(intIndex).CloseWindows
        Set mobjTest(intIndex) = Nothing
    Next
    Call CloseWindows
End Sub

Private Sub imgSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = (Button = 1)
    If Not mblnMove Then Exit Sub
    
    Type_Scale.x = x
    Type_Scale.y = y
End Sub

Private Sub imgSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim dblLeft As Double
    If Not mblnMove Then Exit Sub
    
    dblLeft = imgSplit.Left + x - Type_Scale.x
    If dblLeft < 1000 Or dblLeft > Me.ScaleWidth - 1000 Then Exit Sub
    
    With imgSplit
        .Move .Left + x - Type_Scale.x
    End With
    Call Form_Resize
End Sub

Private Sub imgSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub lblNote_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 0 Then Exit Sub
    mblnMove = (Button = 1)
    If Not mblnMove Then Exit Sub
    
    Type_Scale.x = x
    Type_Scale.y = y
End Sub

Private Sub lblNote_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim dblTop As Double
    If Index = 0 Then Exit Sub
    If Not mblnMove Then Exit Sub
    
    dblTop = lblNote(Index).Top + y - Type_Scale.y
    If dblTop - lblNote(0).Top < 1500 Then Exit Sub
    If Me.ScaleHeight - stbThis.Height - (lblNote(Index).Top + y - Type_Scale.y) - lblNote(1).Height < 1500 Then Exit Sub
    
    With lblNote(Index)
        .Move .Left, .Top + y - Type_Scale.y
    End With
    Call Form_Resize
End Sub

Private Sub lblNote_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub lvwInterface_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwInterface
        .Sorted = False
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwDescending, lvwAscending, lvwDescending)
        .Sorted = True
    End With
End Sub

Private Sub lvwInterface_DblClick()
    Call lvwInterface_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lvwInterface_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lng���� As Long
    Dim IntDO As Integer, intCount As Integer
    Dim arrItem
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    If lvwInterface.ListItems.Count = 0 Then Exit Sub
    If lvwInterface.SelectedItem Is Nothing Then Exit Sub
    
    '��ʾ��ע��ҽ���ӿڵ���ϸ��Ϣ
    lng���� = Mid(Item.Key, 3)
    mnuInterfaceStart.Enabled = (Item.Tag = 1)
    tbrTool.Buttons("Start").Enabled = mnuInterfaceStart.Enabled
    
    '>>ȡ֧�ֵ�ģ��
    Me.lvw����ģ��.ListItems.Clear
    mstrSQL = "Select A.���,A.����,A.˵��" & _
        " From zlPrograms A,zlInsureModuls B" & _
        " Where A.���=B.��� And B.����=" & lng���� & _
        " Order by A.���"
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "ҽ���ӿڹ���")
    With rsTemp
        'ҽ���ӿڻ���ģ�飬Ҳ�����ڵ���
        Do While Not .EOF
            Set lvwItem = Me.lvw����ģ��.ListItems.Add(, "K_" & !���, Nvl(!����), , 1)
            lvwItem.Tag = "zl9Insure"
            .MoveNext
        Loop
        'ҽ���ӿ����ģ��
        Set lvwItem = Me.lvw����ģ��.ListItems.Add(, "K_1111", "����Һ�", , 1)
        lvwItem.Tag = "zl9RegEvent"
        Set lvwItem = Me.lvw����ģ��.ListItems.Add(, "K_1121", "�����շ�", , 1)
        lvwItem.Tag = "zl9OutExse"
        Set lvwItem = Me.lvw����ģ��.ListItems.Add(, "K_1131", "��Ժ�Ǽ�", , 1)
        lvwItem.Tag = "zl9Inpatient"
        Set lvwItem = Me.lvw����ģ��.ListItems.Add(, "K_1132", "���Ժ����", , 1)
        lvwItem.Tag = "zl9Inpatient"
        Set lvwItem = Me.lvw����ģ��.ListItems.Add(, "K_1133", "סԺ����", , 1)
        lvwItem.Tag = "zl9InExse"
        Set lvwItem = Me.lvw����ģ��.ListItems.Add(, "K_1137", "סԺ����", , 1)
        lvwItem.Tag = "zl9InExse"
    End With
    
    '>>ȡ֧�ֿ�˵��
    lvw������.ListItems.Clear
    mstrSQL = "Select �ļ���,˵�� From zlInsureBase Where ����=" & lng����
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "ҽ���ӿڹ���")
    With rsTemp
        Do While Not .EOF
            Call Me.lvw������.ListItems.Add(, "K_" & .AbsolutePosition, Nvl(!�ļ���) & "," & Nvl(!˵��, "��"), , 1)
            .MoveNext
        Loop
    End With
    
    '>>ȡ֧��ҵ��˵��
    mstrSQL = "Select ҵ��,���� From zlInsureOperation Where ����=" & lng����
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "ҽ���ӿڹ���")
    
    With rsTemp
        If .RecordCount <> 0 Then
            For IntDO = 1 To 4
                .Filter = "ҵ��=" & IntDO
                Do While Not .EOF
                    If Trim(Nvl(!����)) <> "" Then
                        arrItem = Split(!����, "|")
                        For intCount = 0 To UBound(arrItem)
                            Call lvw������.ListItems.Add(, "K_" & lvw������.ListItems.Count + 1, arrItem(intCount), , 1)
                        Next
                    End If
                    .MoveNext
                Loop
            Next
        End If
    End With
    
    '���õ�ǰѡ���ҽ��
    SaveSetting "ZLSOFT", "����ȫ��", "�Ƿ�֧��ҽ��", "Yes"
    SaveSetting "ZLSOFT", "����ȫ��", "ҽ�����", lng����
    
    '���ø��ؼ�����ť��״̬
    Call SetEnabled(True)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub lvwInterface_KeyDown(KeyCode As Integer, Shift As Integer)
    If lvwInterface.ListItems.Count = 0 Then Exit Sub
    If lvwInterface.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("��ȷ��Ҫ���ø�ҽ���ӿ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then Exit Sub
    Call mnuInterfaceStart_Click
End Sub

Private Sub lvw����ģ��_DblClick()
    Call lvw����ģ��_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lvw����ģ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intIndex As Integer
    If lvw����ģ��.ListItems.Count = 0 Then Exit Sub
    If lvw����ģ��.SelectedItem Is Nothing Then Exit Sub
    
    Call FindObject
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpTitle_Click()
    '
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuInterfaceInstall_Click()
    If Not InitConfigure Then Exit Sub
    Me.stbThis.Panels(2).Text = "���ڰ�װҽ���ӿ�..."
    If Not mobjConfigure.I_Install(mstrServer) Then
        Me.stbThis.Panels(2).Text = ""
        Exit Sub
    End If
    Call LoadInterface
    MsgBox "ҽ���ӿڲ�����װ�ɹ���", vbInformation, gstrSysname
End Sub

Private Sub mnuInterfaceStart_Click()
    Dim lng���� As Long
    
    If lvwInterface.ListItems.Count = 0 Then Exit Sub
    If lvwInterface.SelectedItem Is Nothing Then Exit Sub
    If lvwInterface.SelectedItem.Tag = 2 Then Exit Sub
    
    lng���� = Mid(lvwInterface.SelectedItem.Key, 3)
    mstrSQL = "ZL_ZLINSURECOMPONENTS_START(" & lng���� & ")"
    gcnOracle.Execute mstrSQL, , adCmdStoredProc
    
    Call LoadInterface
End Sub

Private Sub mnuInterfaceUninstall_Click()
    Dim intInsure As Integer
    If lvwInterface.SelectedItem Is Nothing Then Exit Sub
    intInsure = lvwInterface.SelectedItem
    
    If Not InitConfigure Then Exit Sub
    Me.stbThis.Panels(2).Text = "����ж��ҽ���ӿ�..."
    If Not mobjConfigure.I_UnInstall(intInsure) Then
        Me.stbThis.Panels(2).Text = ""
        Exit Sub
    End If
    
    Me.stbThis.Panels(2).Text = ""
    lvwInterface.ListItems.Remove lvwInterface.SelectedItem.Key
    '����У�����õ���¼�����ʾ��ϸ��Ϣ��������ؿؼ�����ť����Ϊ������ѡ��״̬
    If Me.lvwInterface.ListItems.Count <> 0 Then
        Call lvwInterface_ItemClick(lvwInterface.ListItems(1))
    Else
        Call SetEnabled(False)
        lvw����ģ��.ListItems.Clear
        lvw������.ListItems.Clear
    End If
    
    MsgBox "ҽ���ӿڲ���ж�سɹ���", vbInformation, gstrSysname
End Sub

Private Sub tbrTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Install"
        Call mnuInterfaceInstall_Click
    Case "Uninstall"
        Call mnuInterfaceUninstall_Click
    Case "Start"
        Call mnuInterfaceStart_Click
    Case "Exit"
        Call mnuFileExit_Click
    Case "Help"
        Call mnuHelpTitle_Click
    End Select
End Sub

Private Sub SetEnabled(ByVal BlnState As Boolean)
    mnuInterfaceUninstall.Enabled = BlnState
    tbrTool.Buttons("Uninstall").Enabled = BlnState
End Sub

Private Function RegistFile() As Boolean
    Const strRegist As String = "Regist.txt"
    '���ע���ļ��ĺϷ���
    RegistFile = True
End Function

Private Sub FindObject()
    '�����Ƿ��Ѵ���ָ���Ĳ��������δ�����򴴽�
    Dim strClass As String
    Dim strObject As String
    Dim lngModul As Long
    Dim blnExist As Boolean
    Dim objTest As Object
    Dim intIndex As Integer, intCount As Integer
    Const lngSys As Long = 100
    On Error Resume Next
    
    strObject = UCase(lvw����ģ��.SelectedItem.Tag)
    lngModul = Val(Mid(lvw����ģ��.SelectedItem.Key, 3))
    strClass = strObject & ".cls" & Mid(strObject, 4)
    
    Err = 0
    intCount = UBound(mobjTest)
    If Err <> 0 Then intCount = -1
    
    For intIndex = 0 To intCount
        If mstrTest(intIndex) = UCase(strObject) Then
            blnExist = True
            Exit For
        End If
    Next
    
    '��������
    If blnExist = False Then
        If Not objTest Is Nothing Then
            Call objTest.CloseWindows
            Set objTest = Nothing
        End If
        
        Err = 0
        Set objTest = CreateObject(strClass)
        If Err <> 0 Then
            MsgBox "�޷������ò�������ȷ���Ƿ��Ѱ�װ��", vbInformation
            Exit Sub
        End If
        
        ReDim Preserve mobjTest(intIndex) As Object
        ReDim Preserve mstrTest(intIndex) As String
        Set mobjTest(intIndex) = objTest
        mstrTest(intIndex) = UCase(strObject)
    End If
    
    On Error GoTo ErrHand
    Call mobjTest(intIndex).CodeMan(lngSys, lngModul, gcnOracle, Nothing, "ZLHIS")
    
    Me.WindowState = 1
    Exit Sub
ErrHand:
    MsgBox Err.Description, vbInformation, gstrSysname
End Sub

Private Function IsDBA() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '�жϴ�����û��ǲ��������߻�DBA�û�
    
    mstrSQL = "SELECT 1 FROM DUAL " & _
            " WHERE EXISTS(SELECT 1 FROM ZLSYSTEMS WHERE ������='" & UCase(mstrUser) & "')"
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "�жϸ��û��ǲ��������߻�DBA�û�")
    IsDBA = (rsTemp.RecordCount <> 0)
End Function

Private Function InitConfigure() As Boolean
    If mobjConfigure Is Nothing Then
        On Error Resume Next
        Err = 0
        Set mobjConfigure = CreateObject("zl9I_Configure.clsI_Configure")
        If Err <> 0 Then
            MsgBox "��Ҫ�����ʧ���޷����ҽ���ӿڲ����İ�װ��ж�أ�", vbInformation, gstrSysname
            Exit Function
        End If
        Call mobjConfigure.InitOracle(gcnOracle)
    End If
    
    InitConfigure = True
End Function
