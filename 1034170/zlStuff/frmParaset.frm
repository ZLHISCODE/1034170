VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParaset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "frmParaset.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab tabMain 
      Height          =   7335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "����(&0)"
      TabPicture(0)   =   "frmParaset.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra�ƿ����̿���"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra�ⷿѡ��"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraBidMess"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "����У��(&1)"
      TabPicture(1)   =   "frmParaset.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkProduceDate"
      Tab(1).Control(1)=   "fraCheck"
      Tab(1).Control(2)=   "vsfCheck"
      Tab(1).Control(3)=   "lblComment"
      Tab(1).ControlCount=   4
      Begin VB.CheckBox chkProduceDate 
         Caption         =   "�������ڴ���ע��֤Ч�ڼ��"
         Height          =   255
         Left            =   -74760
         TabIndex        =   43
         Top             =   5160
         Width           =   2775
      End
      Begin VB.Frame fraBidMess 
         Caption         =   "��ⵥ�۳��б�ɱ���"
         Height          =   735
         Left            =   120
         TabIndex        =   34
         Top             =   4860
         Width           =   3675
         Begin VB.OptionButton optBidMess 
            Caption         =   "��ֹ"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optBidMess 
            Caption         =   "��ʾ"
            Height          =   180
            Index           =   1
            Left            =   1200
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optBidMess 
            Caption         =   "������"
            Height          =   180
            Index           =   2
            Left            =   2280
            TabIndex        =   35
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame fraCheck 
         Caption         =   "ѡ��У�鷽ʽ"
         Height          =   615
         Left            =   -74760
         TabIndex        =   30
         Top             =   5520
         Width           =   7095
         Begin VB.OptionButton optCheck 
            Caption         =   "У��δͨ��ʱ����"
            Height          =   180
            Index           =   1
            Left            =   3360
            TabIndex        =   32
            Top             =   280
            Width           =   2175
         End
         Begin VB.OptionButton optCheck 
            Caption         =   "У��δͨ��ʱ��ֹ����"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   31
            Top             =   280
            Width           =   2175
         End
      End
      Begin VB.Frame fra�ⷿѡ�� 
         Caption         =   "�ⷿѡ��"
         Height          =   1665
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   3675
         Begin VB.CheckBox chkStock 
            Caption         =   "����ѡ��ⷿ"
            Height          =   375
            Left            =   210
            TabIndex        =   27
            Top             =   240
            Width           =   2805
         End
         Begin VB.Label lbl�ⷿѡ��˵�� 
            Caption         =   "    ���ѡ��ⷿ�����ڵ�������'���пⷿ'Ȩ���˾Ϳ���ѡ��ͬ�ⷿ�����򣬲���ѡ��ⷿ��"
            Height          =   615
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   3285
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "���ĵ�λ"
         Height          =   1665
         Left            =   3840
         TabIndex        =   20
         Top             =   480
         Width           =   3675
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   390
            Width           =   2655
         End
         Begin VB.ComboBox CboUnit1 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   780
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "    ��ѡ��һ���������ϵ�λ���ڵ��������У������������Ͻ������ֵ�λ��"
            Height          =   405
            Left            =   240
            TabIndex        =   25
            Top             =   1170
            Width           =   3315
         End
         Begin VB.Label lbl�̵�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�̵��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   24
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl�̵㵥 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�̵㵥"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   23
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "����"
         Height          =   3345
         Left            =   3840
         TabIndex        =   11
         Top             =   2250
         Width           =   3675
         Begin VB.CheckBox chk�ӳ���� 
            Caption         =   "ʱ�������Լӳ������"
            Height          =   255
            Left            =   390
            TabIndex        =   46
            Top             =   2040
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chkȡ�ϴ��ۼ� 
            Caption         =   "ʱ���������ʱȡ�ϴ��ۼ�"
            Height          =   255
            Left            =   390
            TabIndex        =   45
            Top             =   2280
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk�ֶμӳ���� 
            Caption         =   "ʱ�����İ��ֶμӳ����"
            Height          =   255
            Left            =   390
            TabIndex        =   44
            Top             =   2520
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chkʱ�۵��� 
            Caption         =   "ʱ�����İ����ε���"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   390
            TabIndex        =   41
            Top             =   360
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.CheckBox chkSet�ⷿ 
            Caption         =   "�����̵�û�����ô洢�ⷿ������"
            Height          =   255
            Left            =   390
            TabIndex        =   39
            Top             =   1230
            Visible         =   0   'False
            Width           =   3105
         End
         Begin VB.CheckBox chk��ֵ����¼�� 
            Caption         =   "��ֵ���ı�����д��ϸ��Ϣ"
            Height          =   255
            Left            =   390
            TabIndex        =   13
            Top             =   1755
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk��ҩ������� 
            Caption         =   "���������ǰ��Ҫ���в���˲�"
            Height          =   255
            Left            =   390
            TabIndex        =   33
            Top             =   1235
            Visible         =   0   'False
            Width           =   3180
         End
         Begin VB.CheckBox chkSavePrint 
            Caption         =   "���ݴ��̺��Զ���ӡ"
            Height          =   255
            Left            =   390
            TabIndex        =   19
            Top             =   455
            Width           =   1935
         End
         Begin VB.CheckBox chkVerifyPrint 
            Caption         =   "������˺��Զ���ӡ"
            Height          =   255
            Left            =   390
            TabIndex        =   18
            Top             =   715
            Width           =   1935
         End
         Begin VB.CheckBox chk�޸ĵ��ݺ� 
            Caption         =   "�����޸ĵ��ݺ�"
            Height          =   255
            Left            =   390
            TabIndex        =   17
            Top             =   972
            Visible         =   0   'False
            Width           =   2505
         End
         Begin VB.CheckBox chkFixPrice 
            Caption         =   "�⹺��ⶨ�۲ɹ�"
            Height          =   255
            Left            =   390
            TabIndex        =   16
            Top             =   195
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.CheckBox chk�����޸������� 
            Caption         =   "�����޸�������"
            Height          =   255
            Left            =   390
            TabIndex        =   15
            Top             =   1235
            Visible         =   0   'False
            Width           =   2700
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "���ݴ�ӡ����(&S)"
            Height          =   350
            Left            =   360
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2880
            Width           =   2925
         End
         Begin VB.CheckBox chk�б����� 
            Caption         =   "�б����Ŀ�ѡ����б굥λ���"
            Height          =   255
            Left            =   390
            TabIndex        =   12
            Top             =   1495
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk�������� 
            Caption         =   "���������""��������""���Ե����Ľ�������"
            Height          =   360
            Left            =   390
            TabIndex        =   38
            Top             =   1550
            Width           =   2850
         End
      End
      Begin VB.Frame fra���� 
         Caption         =   "����ʽ"
         Height          =   2505
         Left            =   120
         TabIndex        =   7
         Top             =   2250
         Width           =   3675
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   390
            Width           =   2415
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   390
            Width           =   885
         End
         Begin VB.Label lbl����˵�� 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "    �����������ã���Ӱ�����б༭�����е��ݵ���ʾ���ݵ�����ʽ��ȱʡ�����û������˳����ʾ�����ݵ�����"
            ForeColor       =   &H80000008&
            Height          =   825
            Left            =   180
            TabIndex        =   10
            Top             =   930
            Width           =   3345
         End
      End
      Begin VB.Frame fra�ƿ����̿��� 
         Caption         =   "�ƿ����̿���"
         Height          =   1275
         Left            =   120
         TabIndex        =   4
         Top             =   5880
         Width           =   7365
         Begin VB.CheckBox chkRequestStrike 
            Caption         =   "�ƿ����ʱ������ⷿ��Ҫ���������"
            Height          =   180
            Left            =   180
            TabIndex        =   40
            Top             =   960
            Value           =   1  'Checked
            Width           =   3705
         End
         Begin VB.CheckBox chk�ƿ����̿��� 
            Caption         =   "�ƿ�ʱ��Ҫ���ϡ����͡�������һ���̡�"
            Height          =   180
            Left            =   180
            TabIndex        =   5
            Top             =   270
            Value           =   1  'Checked
            Width           =   6945
         End
         Begin VB.Label lbl�ƿ�˵�� 
            Caption         =   "ע�⣺������򹴣���ô����д�ƿⵥ������һ����˲�������˺��Զ���ɱ��ϡ����͡�������һ���̡����ǰ�����޸ĵ��ݡ�"
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   6945
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfCheck 
         Height          =   4125
         Left            =   -74760
         TabIndex        =   42
         Top             =   960
         Width           =   7095
         _cx             =   12515
         _cy             =   7276
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
         BackColorSel    =   16711680
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   25
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmParaset.frx":0044
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
         VirtualData     =   0   'False
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
      Begin VB.Label lblComment 
         Caption         =   $"frmParaset.frx":02EF
         Height          =   540
         Left            =   -74760
         TabIndex        =   29
         Top             =   480
         Width           =   7140
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   360
      TabIndex        =   2
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5160
      TabIndex        =   0
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6360
      TabIndex        =   1
      Top             =   7680
      Width           =   1100
   End
End
Attribute VB_Name = "frmParaset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Private mlngModule As Long '
Private mstrPrivs As String '
Private mblnHavePriv As Boolean
Private mblnFirstLoad As Boolean    '��¼�Ƿ��һ�μ���

Private Sub chk�ֶμӳ����_Click()
    If chk�ֶμӳ����.Value = 1 Then
        chk�ӳ����.Value = 0
        chkȡ�ϴ��ۼ�.Value = 0
    End If
End Sub

Private Sub chk�ӳ����_Click()
    If chk�ӳ����.Value = 1 Then
        chkȡ�ϴ��ۼ�.Value = 0
        chk�ֶμӳ����.Value = 0
    End If
End Sub

Private Sub chkȡ�ϴ��ۼ�_Click()
    If chkȡ�ϴ��ۼ�.Value = 1 Then
        chk�ӳ����.Value = 0
        chk�ֶμӳ����.Value = 0
    End If
End Sub
Private Function ISValid() As Boolean
    Dim i As Integer
    Dim blnAllUnCheck As Boolean
    
    '����У��
    If tabMain.TabVisible(1) = True Then
        blnAllUnCheck = True
        With vsfCheck
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("У��")) <> "" Then
                    blnAllUnCheck = False
                    Exit For
                End If
            Next
        End With
        
        '���ѡ����У����Ŀ�������ѡ��У�鷽ʽ
        If blnAllUnCheck = False And optCheck(0).Value = 0 And optCheck(1).Value = 0 Then
            MsgBox "��ѡ������У�鷽ʽ��", vbExclamation, gstrSysName
            tabMain.Tab = 1
            If vsfCheck.Enabled Then vsfCheck.SetFocus
            Exit Function
        End If
    End If
    
    ISValid = True
End Function

Private Sub Save����У��()
    Dim i As Integer
    Dim strCheck As String
    Dim blnAllUnCheck As Boolean
    
    If mstrFunction = "�����⹺������" Or mstrFunction = "���ļƻ�����" Then
        blnAllUnCheck = True
        
        '��������У����Ŀ�ͷ�ʽ����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....
        With vsfCheck
            For i = 1 To .Rows - 1
                strCheck = IIf(strCheck = "", "", strCheck & ";") & .TextMatrix(i, .ColIndex("���")) & "," & .TextMatrix(i, .ColIndex("У����Ŀ")) & "," & _
                    IIf(.TextMatrix(i, .ColIndex("У��")) = "", 0, 1)
                    
                If .TextMatrix(i, .ColIndex("У��")) <> "" Then blnAllUnCheck = False
            Next
        End With
        
        If blnAllUnCheck = True Then
            strCheck = "0|" & strCheck
        ElseIf optCheck(0).Value = True Then
            strCheck = "2|" & strCheck
        Else
            strCheck = "1|" & strCheck
        End If
            
        Call zlDatabase.SetPara("����У��", strCheck, glngSys, mlngModule)
    End If
    
End Sub

Private Sub Load����У��()
    Dim i As Integer
    Dim n As Integer
    Dim strCheck As String
    Dim intCheckType As Integer
    Dim arrColumn
    
    On Error Resume Next
    
    If mstrFunction = "�����⹺������" Or mstrFunction = "���ļƻ�����" Then
        '����У����Ŀ�ͷ�ʽ�ı����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....
        strCheck = zlDatabase.GetPara("����У��", glngSys, mlngModule, "", Array(vsfCheck, fraCheck), mblnHavePriv)
        
        If strCheck <> "" Then
            If mstrFunction = "�����⹺������" Then
                chkProduceDate.Value = IIf(Val(zlDatabase.GetPara("��������Ч�ڼ��", glngSys, mlngModule, "0", Array(chkProduceDate), mblnHavePriv)) = 1, 1, 0)
            End If
            
            If InStr(1, strCheck, "|") > 0 Then
                'У�鷽ʽ��0-����飻1�����ѣ�2����ֹ
                intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
                If intCheckType = 2 Then
                    optCheck(0).Value = True
                ElseIf intCheckType = 1 Then
                    optCheck(1).Value = True
                End If
                
                strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)
                 
                If strCheck <> "" Then
                    strCheck = strCheck & ";"
                    arrColumn = Split(strCheck, ";")
                    For n = 0 To UBound(arrColumn)
                        If arrColumn(n) <> "" Then
                            With vsfCheck
                                For i = 1 To .Rows - 1
                                    If Split(arrColumn(n), ",")(0) = .TextMatrix(i, .ColIndex("���")) And Split(arrColumn(n), ",")(1) = .TextMatrix(i, .ColIndex("У����Ŀ")) Then
                                        If Val(Split(arrColumn(n), ",")(2)) = 1 Then
                                            .TextMatrix(i, .ColIndex("У��")) = "��"
                                        End If
                                    End If
                                Next
                            End With
                        End If
                    Next
                End If
            End If
        End If
    End If
End Sub
Private Sub Cbo����_Click()
    If cbo����.ListCount < 1 Then Exit Sub
    cbo����.Enabled = Not (cbo����.ListIndex = 0)
    If Not cbo����.Enabled Then cbo����.ListIndex = 0
End Sub
Private Sub chk�����ۼ�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkRequestStrike_Click()
    '����Ϊ����Ҫ����ʱ��Ҫ����Ƿ���δ��˵ĳ������뵥����������ܸı�
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If mblnFirstLoad = True Then
        If chkRequestStrike.Value = 0 Then
            If MsgBox("��������Ƿ����δ��˵ĳ������뵥��������Ҫ�ϳ�ʱ�䣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                '�ù�����10.34�汾����������һ�������������ڷ�Χ������ȫ��ɨ�裬��˿��Ǵ�34�汾�޸����ڿ�ʼ
                gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ���� = 19 And Mod(��¼״̬, 3) = 2 And ������� Is Null " & _
                    " And �������� Between To_Date('2014/2/20 00:00:00', 'yyyy-mm-dd hh24:mi:ss') And Sysdate And Rownum = 1"
                
                DoEvents
                zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���δ��˵ĳ������뵥")
                
                DoEvents
                zlCommFun.StopFlash
                
                If rsTemp.RecordCount > 0 Then
                    MsgBox "����δ��˵ĳ������뵥�����ܸı�˲�����", vbInformation, gstrSysName
                    chkRequestStrike.Value = 1
                End If
            Else
                chkRequestStrike.Value = 1
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub
Private Sub SetCtlEnabled()
    '-----------------------------------------------------------------------------
    '����:Ȩ������
    '-----------------------------------------------------------------------------
    Dim blnPara As Boolean
    blnPara = InStr(1, mstrPrivs, ";��������;") > 0
    
    chkFixPrice.Enabled = blnPara
    chk�޸ĵ��ݺ�.Enabled = blnPara
    chk�����޸�������.Enabled = blnPara
    chk�б�����.Enabled = blnPara
    chk�ƿ����̿���.Enabled = blnPara
    chk��ֵ����¼��.Enabled = blnPara
    chk��������.Enabled = blnPara
    
    If tabMain.TabVisible(1) = True Then
        vsfCheck.Enabled = blnPara
        fraCheck.Enabled = blnPara
    End If
    If mlngModule = 1726 Then
        fra�ⷿѡ��.Enabled = False
    End If
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '����:�����ݿⱣ���������
    '����:����ɹ�����True,���򷵻�False
    '����:���˺�
    '����:2007/12/24
    '------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:
    gcnOracle.BeginTrans
    
    Call zlDatabase.SetPara("��ⵥ�۳��б굥��", IIf(optBidMess(0).Value, 0, IIf(optBidMess(1).Value, 1, 2)), glngSys, mlngModule, IIf(fraBidMess.Enabled, True, False))
    Call zlDatabase.SetPara("��������", CStr(cbo����.ListIndex) & CStr(cbo����.ListIndex), glngSys, mlngModule)   '
    Call zlDatabase.SetPara("���̴�ӡ", IIf(chkSavePrint.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("��˴�ӡ", IIf(chkVerifyPrint.Value = 1, 1, 0), glngSys, mlngModule)
    
    If chk��ֵ����¼��.Visible Then
        Call zlDatabase.SetPara("��ֵ���ı�����д��ϸ��Ϣ", IIf(chk��ֵ����¼��.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    
    If mlngModule = 1725 Then
        Call zlDatabase.SetPara("�Ƿ�ѡ����", IIf(chkStock.Value = 1, 1, 0), glngSys, mlngModule)
    Else
        Call zlDatabase.SetPara("�Ƿ�ѡ��ⷿ", IIf(chkStock.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    Call zlDatabase.SetPara(IIf(mlngModule = 1719, "�̵��λ", "���ĵ�λ"), cboUnit.ListIndex, glngSys, mlngModule)
    Call zlDatabase.SetPara("�޸ĵ��ݺ�", IIf(chk�޸ĵ��ݺ�.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("��������", IIf(chk��������.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("�洢�ⷿ", IIf(chkSet�ⷿ.Value = 1, 1, 0), glngSys, mlngModule)
    
    If chkFixPrice.Visible = True Then
        Select Case mstrFunction
            Case "�����⹺������"
                Call zlDatabase.SetPara("���۲ɹ�", IIf(chkFixPrice.Value = 1, 1, 0), glngSys, mlngModule)
                Call zlDatabase.SetPara("�޸Ĳɹ��޼�", IIf(chk�����޸�������.Value = 1, 1, 0), glngSys, mlngModule)
                Call zlDatabase.SetPara("�б����Ŀ�ѡ����б굥λ���", IIf(chk�б�����.Value = 1, 1, 0), glngSys, mlngModule)
'                Call zlDatabase.SetPara("У�鹩Ӧ������", IIf(chk��Ӧ��У��.Value = 1, 1, 0), glngSys, mlngModule)
            Case ""
        End Select
    End If
    If CboUnit1.Visible Then
        Call zlDatabase.SetPara("��¼����λ", CboUnit1.ListIndex, glngSys, mlngModule)
    End If
    If fra�ƿ����̿���.Visible = True Then
        Call zlDatabase.SetPara("�ƿ�����", IIf(chk�ƿ����̿���.Value = 1, 1, 0), glngSys, mlngModule)
        Call zlDatabase.SetPara("��������", IIf(chkRequestStrike.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    
    If mlngModule = 1717 Then
        Call zlDatabase.SetPara("�������", IIf(chk��ҩ�������.Value = 1, 1, 0), glngSys, mlngModule)
    End If
    
    '���ĵ��۹���
    If mlngModule = 1726 Then
        zlDatabase.SetPara "ʱ�����İ����ε���", chkʱ�۵���.Value, glngSys, mlngModule
    End If
    
    Save����У��
    If mlngModule = 1712 Then
        '�����⹺��������������Ч�ڼ��
        zlDatabase.SetPara "��������Ч�ڼ��", chkProduceDate.Value, glngSys, mlngModule
        
        zlDatabase.SetPara "ʱ�������ԼӼ������", chk�ӳ����.Value, glngSys, mlngModule
        zlDatabase.SetPara "ʱ���������ʱȡ�ϴ��ۼ�", chkȡ�ϴ��ۼ�.Value, glngSys, mlngModule
        zlDatabase.SetPara "���ķֶμӳ���", chk�ֶμӳ����.Value, glngSys, mlngModule
    End If
    
    '�������
    If mlngModule = 1714 Then
        zlDatabase.SetPara "ʱ�������ԼӼ������", chk�ӳ����.Value, glngSys, mlngModule
        zlDatabase.SetPara "ʱ���������ʱȡ�ϴ��ۼ�", chkȡ�ϴ��ۼ�.Value, glngSys, mlngModule
        zlDatabase.SetPara "���ķֶμӳ���", chk�ֶμӳ����.Value, glngSys, mlngModule
    End If
    
    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Sub cmdOk_Click()
    If ISValid = False Then Exit Sub
    If SaveSet = False Then Exit Sub
    Unload Me
End Sub

Private Sub initPara()
    '-------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:
    '����:���˺�
    '�޸�:2007/12/24
    '-------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim strBidMess As String
    Dim int�ӳ������ As Integer    '��������Ҳ���������
    Dim intȡ�ϴ��ۼ� As Integer    '��������Ҳ���������
    Dim int�ֶμӳ���� As Integer  '��������Ҳ���������
    
    'װ��ȱʡ����
    With cbo����
        .Clear
        .AddItem "����˳��"
        .ItemData(.NewIndex) = 0
        .AddItem "����"
        .ItemData(.NewIndex) = 1
        .AddItem "��������"
        .ItemData(.NewIndex) = 2
        If mstrFunction = "�����̵����" Then
            .AddItem "�ⷿ��λ"
            .ItemData(.NewIndex) = 3
        End If
        .ListIndex = 0
    End With
    
    With cbo����
        .Clear
        .AddItem "����"
        .ItemData(.NewIndex) = 0
        .AddItem "����"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    
    fraBidMess.Visible = False
    
    strValue = zlDatabase.GetPara("��������", glngSys, mlngModule, "00", Array(cbo����, cbo����, fra����, lbl����˵��), mblnHavePriv)
    strValue = IIf(strValue = "", "00", strValue)
    cbo����.ListIndex = Val(Mid(strValue, 1, 1))
    cbo����.ListIndex = Val(Right(strValue, 1))
    cbo����.Enabled = Not (cbo����.ListIndex = 0)
    
    chkSavePrint.Value = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0", Array(chkSavePrint), mblnHavePriv)) = 1, 1, 0)
    chkVerifyPrint.Value = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0", Array(chkVerifyPrint), mblnHavePriv)) = 1, 1, 0)
    If mlngModule = 1725 Then
        chkStock.Value = IIf(Val(zlDatabase.GetPara("�Ƿ�ѡ����", glngSys, mlngModule, "0", Array(chkStock, fra�ⷿѡ��, lbl�ⷿѡ��˵��), mblnHavePriv)) = 1, 1, 0)
    Else
        chkStock.Value = IIf(Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModule, "0", Array(chkStock, fra�ⷿѡ��, lbl�ⷿѡ��˵��), mblnHavePriv)) = 1, 1, 0)
    End If
    
    With CboUnit1
        .Clear
        .AddItem "ɢװ��λ"
        .AddItem "��װ��λ"
    End With

    With cboUnit
        .Clear
        .AddItem "ɢװ��λ"
        .AddItem "��װ��λ"
    End With
    cboUnit.ListIndex = IIf(Val(zlDatabase.GetPara(IIf(mlngModule = 1719, "�̵��λ", "���ĵ�λ"), glngSys, mlngModule, "0", Array(cboUnit, lbl�̵��), mblnHavePriv)) = 1, 1, 0)
    If mstrFunction <> "�����̵����" Then
        CboUnit1.Visible = False
        lbl�̵��.Visible = False
        lbl�̵㵥.Visible = False
        cboUnit.Left = lbl�̵��.Left
        cboUnit.Width = Frame2.Width - cboUnit.Left - 250
        Label2.Top = lbl�̵㵥.Top
    Else
        CboUnit1.ListIndex = IIf(Val(zlDatabase.GetPara("��¼����λ", glngSys, mlngModule, "0", Array(CboUnit1, lbl�̵㵥), mblnHavePriv)) = 1, 1, 0)
    End If
     
     
    chk�޸ĵ��ݺ�.Visible = True
    chk�޸ĵ��ݺ�.Value = IIf(Val(zlDatabase.GetPara("�޸ĵ��ݺ�", glngSys, mlngModule, "0", Array(chk�޸ĵ��ݺ�), mblnHavePriv)) = 1, 1, 0)
    chk��������.Visible = False
    chkSet�ⷿ.Visible = False
    
    Select Case mstrFunction
        Case "�����̵����"
            chkSet�ⷿ.Visible = True
            chkSet�ⷿ.Value = IIf(Val(zlDatabase.GetPara("�洢�ⷿ", glngSys, mlngModule, "0", Array(chkSet�ⷿ), mblnHavePriv)) = 1, 1, 0)
        Case "�����⹺������"
            chkFixPrice.Visible = True
            chk�����޸�������.Visible = True
            chk�б�����.Visible = True
            chk��ֵ����¼��.Visible = True
            fraBidMess.Visible = True

            chkFixPrice.Value = IIf(Val(zlDatabase.GetPara("���۲ɹ�", glngSys, mlngModule, "0", Array(chkFixPrice), mblnHavePriv)) = 1, 1, 0)
            chk�����޸�������.Value = IIf(Val(zlDatabase.GetPara("�޸Ĳɹ��޼�", glngSys, mlngModule, "0", Array(chk�����޸�������), mblnHavePriv)) = 1, 1, 0)
            chk�б�����.Value = IIf(Val(zlDatabase.GetPara("�б����Ŀ�ѡ����б굥λ���", glngSys, mlngModule, "0", Array(chk�б�����), mblnHavePriv)) = 1, 1, 0)
            chk��ֵ����¼��.Value = IIf(Val(zlDatabase.GetPara("��ֵ���ı�����д��ϸ��Ϣ", glngSys, mlngModule, "0", Array(chk��ֵ����¼��), mblnHavePriv)) = 1, 1, 0)
            
            strBidMess = zlDatabase.GetPara("��ⵥ�۳��б굥��", glngSys, mlngModule, , Array(optBidMess(0), optBidMess(1), optBidMess(2), fraBidMess), mblnHavePriv)
            optBidMess(Val(strBidMess)).Value = True
            
            int�ӳ������ = Val(zlDatabase.GetPara("ʱ�������ԼӼ������", glngSys, mlngModule, 1, Array(chk�ӳ����), mblnHavePriv))
            intȡ�ϴ��ۼ� = Val(zlDatabase.GetPara("ʱ���������ʱȡ�ϴ��ۼ�", glngSys, mlngModule, 0, Array(chkȡ�ϴ��ۼ�), mblnHavePriv))
            int�ֶμӳ���� = Val(zlDatabase.GetPara("���ķֶμӳ���", glngSys, mlngModule, 0, Array(chk�ֶμӳ����), mblnHavePriv))
            
            '����������
            If int�ӳ������ = 1 Then
                intȡ�ϴ��ۼ� = 0
                int�ֶμӳ���� = 0
            ElseIf intȡ�ϴ��ۼ� = 1 Then
                int�ӳ������ = 0
                int�ֶμӳ���� = 0
            ElseIf int�ֶμӳ���� = 1 Then
                int�ӳ������ = 0
                intȡ�ϴ��ۼ� = 0
            End If
            
            chk�ӳ����.Visible = True
            chkȡ�ϴ��ۼ�.Visible = True
            chk�ֶμӳ����.Visible = True
            
            chk�ӳ����.Value = int�ӳ������
            chkȡ�ϴ��ۼ�.Value = intȡ�ϴ��ۼ�
            chk�ֶμӳ����.Value = int�ֶμӳ����
        Case "��������������"
            int�ӳ������ = Val(zlDatabase.GetPara("ʱ�������ԼӼ������", glngSys, mlngModule, 1, Array(chk�ӳ����), mblnHavePriv))
            intȡ�ϴ��ۼ� = Val(zlDatabase.GetPara("ʱ���������ʱȡ�ϴ��ۼ�", glngSys, mlngModule, 0, Array(chkȡ�ϴ��ۼ�), mblnHavePriv))
            int�ֶμӳ���� = Val(zlDatabase.GetPara("���ķֶμӳ���", glngSys, mlngModule, 0, Array(chk�ֶμӳ����), mblnHavePriv))
            
            '����������
            If int�ӳ������ = 1 Then
                intȡ�ϴ��ۼ� = 0
                int�ֶμӳ���� = 0
            ElseIf intȡ�ϴ��ۼ� = 1 Then
                int�ӳ������ = 0
                int�ֶμӳ���� = 0
            ElseIf int�ֶμӳ���� = 1 Then
                int�ӳ������ = 0
                intȡ�ϴ��ۼ� = 0
            End If
            
            chk�ӳ����.Visible = True
            chkȡ�ϴ��ۼ�.Visible = True
            chk�ֶμӳ����.Visible = True
            
            chk�ӳ����.Value = int�ӳ������
            chkȡ�ϴ��ۼ�.Value = intȡ�ϴ��ۼ�
            chk�ֶμӳ����.Value = int�ֶμӳ����
        Case "���ļƻ�����", "�����깺����"
            chk�޸ĵ��ݺ�.Visible = False
        Case "�������ù���"
            chk��ҩ�������.Visible = True
            chk��ҩ�������.Value = IIf(Val(zlDatabase.GetPara("�������", glngSys, mlngModule, "0", Array(chk��ҩ�������), mblnHavePriv)) = 1, 1, 0)
            chk��������.Visible = True
            chk��������.Value = IIf(Val(zlDatabase.GetPara("��������", glngSys, mlngModule, "0", Array(chk��������), mblnHavePriv)) = 1, 1, 0)
        Case Else
    End Select
    
    If mstrFunction <> "�����⹺������" Then
        fra����.Height = Frame3.Height
    End If
    
    fra����.Enabled = (InStr(1, "���ĸ�������", mstrFunction) = 0)
    If fra����.Enabled = False Then
        cbo����.Enabled = False
        cbo����.Enabled = False
    End If
    If Frame2.Enabled = False Then
        cboUnit.Enabled = False
    End If
    
    fra�ⷿѡ��.Enabled = (InStr(1, "���ĸ�������", mstrFunction) = 0)
    Me.cmdPrintSet.Enabled = InStr(1, gstrPrivs, ";���ݴ�ӡ;") <> 0
    
    If fra�ⷿѡ��.Enabled = False Then
        chkStock.Enabled = False
    End If
    
    If mstrFunction = "�����ƿ����" Then
        mblnFirstLoad = False
        chk�ƿ����̿���.Value = IIf(Val(zlDatabase.GetPara("�ƿ�����", glngSys, mlngModule, "0", Array(chk�ƿ����̿���, lbl�ƿ�˵��, fra�ƿ����̿���), mblnHavePriv)) = 1, 1, 0)
        chkRequestStrike.Value = IIf(Val(zlDatabase.GetPara("��������", glngSys, mlngModule, "0", Array(chkRequestStrike, fra�ƿ����̿���), mblnHavePriv)) = 1, 1, 0)
        mblnFirstLoad = True
    Else
        fra�ƿ����̿���.Visible = False
        
        tabMain.Height = tabMain.Height - fra�ƿ����̿���.Height
        
        cmdHelp.Top = cmdHelp.Top - fra�ƿ����̿���.Height
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        Me.Height = Me.Height - fra�ƿ����̿���.Height
    End If
    
    '����У��ҳ��
    tabMain.TabVisible(1) = mstrFunction = "�����⹺������" Or mstrFunction = "���ļƻ�����"
    If tabMain.TabVisible(1) = True Then
        With vsfCheck
            .MergeCol(0) = True
            .MergeCells = flexMergeRestrictColumns
        End With
        If mstrFunction = "�����⹺������" Then
            fraCheck.Top = tabMain.Height - fraCheck.Height - 100
            chkProduceDate.Top = fraCheck.Top - chkProduceDate.Height - 100
            vsfCheck.Height = chkProduceDate.Top - vsfCheck.Top - 100
        Else
            '���ļƻ����ڲ��������������� ���Բ���Ҫ�˲���
            fraCheck.Top = tabMain.Height - fraCheck.Height - 100
            vsfCheck.Height = fraCheck.Top - vsfCheck.Top - 100
        End If
        
        If mstrFunction = "�����⹺������" Then
            lblComment.Caption = "    ˵���������⹺���༭����ʱ�Ƿ�У�����ġ������̡���Ӧ�̵���Ϣ�Ƿ��������������Ƿ���ڡ���ѡ����Ҫ����У�����Ŀ����˫����У�顱�д򹴡�"
        ElseIf mstrFunction = "���ļƻ�����" Then
            lblComment.Caption = "    ˵�������ļƻ�������˵���ʱ�Ƿ�У�����ġ������̡���Ӧ�̵���Ϣ�Ƿ��������������Ƿ���ڡ���ѡ����Ҫ����У�����Ŀ����˫����У�顱�д򹴡�"
        End If
        
        Load����У��
    End If
    
    chkʱ�۵���.Visible = False
    If mstrFunction = "���ĵ��۹���" Then
        chkʱ�۵���.Value = IIf(Val(zlDatabase.GetPara("ʱ�����İ����ε���", glngSys, mlngModule, "0", Array(chkʱ�۵���), mblnHavePriv)) = 1, 1, 0)
        fra�ⷿѡ��.Visible = False
        fra����.Visible = False
        fraBidMess.Visible = False
        Frame3.Visible = True
        Frame3.Top = Frame2.Top
        Frame3.Height = Frame2.Height
        Frame3.Enabled = True
        chkʱ�۵���.Visible = True
        fra�ƿ����̿���.Visible = False
        Frame2.Left = fra�ⷿѡ��.Left
        chkSavePrint.Visible = False
        chkVerifyPrint.Visible = False
        chk�޸ĵ��ݺ�.Visible = False
        tabMain.Height = Frame2.Top + Frame2.Height + 100
        tabMain.Width = Frame2.Width + Frame3.Width + 300
        cmdOK.Top = tabMain.Height + tabMain.Top + 150
        cmdHelp.Top = cmdOK.Top
        cmdHelp.Left = 100
        cmdCancel.Top = cmdOK.Top
        cmdCancel.Left = tabMain.Width - cmdCancel.Width
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
        Me.Height = tabMain.Top + tabMain.Height + 1100
        Me.Width = tabMain.Width + 250
    End If
End Sub

Public Sub ���ò���(ByVal lngModule As Long, ByVal strPrivs As String, ByVal frmMain As Form, Optional ByVal strFunction As String = "")
    '-------------------------------------------------------------------------------------------------------------
    '����:������ص��ݲ����Ŀ��Ʋ���
    '����:lngModule-ģ���
    '     strȨ�޴�-Ȩ�޴�
    '     frmMain-���õ�������
    '     strFunction-����˵��
    '����:
    '����:���˺�
    '�޸�:2007/12/24
    '-------------------------------------------------------------------------------------------------------------
    mstrPrivs = strPrivs: mlngModule = lngModule: mstrFunction = strFunction
    mblnHavePriv = IsHavePrivs(mstrPrivs, "��������")
    
    Call initPara
    Call SetCtlEnabled
    frmParaset.Show vbModal, frmMain
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    strBill = "ZL1_BILL_" & glngModul
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFirstLoad = False
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        If vsfCheck.Enabled = True Then vsfCheck.SetFocus
        If vsfCheck.TextMatrix(4, 1) = "" Then
            chkProduceDate.Enabled = False
            chkProduceDate.Value = 0
        Else
            chkProduceDate.Enabled = True
        End If
    End If
End Sub

Private Sub vsfCheck_DblClick()
    With vsfCheck
        If .Row = 0 Then Exit Sub
        If .Col <> .ColIndex("У��") Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .TextMatrix(.Row, .Col) = "��" Then
            .TextMatrix(.Row, .Col) = ""
            If .Row = 4 Then
                'ע��֤��Ч�ڼ��
                chkProduceDate.Enabled = False
                chkProduceDate.Value = 0
            End If
        Else
            .TextMatrix(.Row, .Col) = "��"
            If .Row = 4 Then
                'ע��֤��Ч�ڼ��
                chkProduceDate.Enabled = True
            End If
        End If
    End With
End Sub


