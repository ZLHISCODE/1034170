VERSION 5.00
Begin VB.Form frmReportSetup 
   BorderStyle     =   0  'None
   ClientHeight    =   6795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraReportSetup 
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.Frame fra 
         Height          =   1485
         Index           =   5
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   7695
         Begin VB.Frame fra 
            Height          =   1125
            Index           =   9
            Left            =   480
            TabIndex        =   47
            Top             =   270
            Width           =   2055
            Begin VB.CheckBox chkIgnorePosi 
               Caption         =   "���Խ����������"
               Height          =   180
               Left            =   120
               TabIndex        =   50
               ToolTipText     =   "����¼�ʹ��������ԡ�"
               Top             =   0
               Width           =   1800
            End
            Begin VB.CheckBox chkReportAfterResult 
               Caption         =   "���������Ϊ����"
               Height          =   180
               Left            =   120
               TabIndex        =   49
               ToolTipText     =   "��д����ʱ��û��¼����ϣ���Ĭ�ϼ�¼Ϊ���ԡ�"
               Top             =   720
               Width           =   1740
            End
            Begin VB.CheckBox chkDefaultPosi 
               Caption         =   "��Ͻ��Ĭ������"
               Height          =   300
               Left            =   120
               TabIndex        =   48
               ToolTipText     =   "����������ѡ�񴰿ڣ�Ĭ��ѡ�����ԡ�"
               Top             =   300
               Width           =   1815
            End
         End
         Begin VB.CheckBox chkConformDetermine 
            Caption         =   "��������ж�"
            Height          =   180
            Left            =   2640
            TabIndex        =   46
            ToolTipText     =   "�������������ܺͲ˵�"
            Top             =   280
            Width           =   1455
         End
         Begin VB.CheckBox chkReportLevel 
            Caption         =   "���������ȼ�"
            Height          =   180
            Left            =   2640
            TabIndex        =   45
            Top             =   657
            Width           =   1410
         End
         Begin VB.CheckBox chkImageLevel 
            Caption         =   "Ӱ�������ȼ�"
            Height          =   180
            Left            =   2640
            TabIndex        =   44
            Top             =   1035
            Width           =   1410
         End
         Begin VB.TextBox txtReportLevel 
            Height          =   270
            Left            =   4050
            TabIndex        =   43
            Text            =   "��,��"
            Top             =   600
            Width           =   1035
         End
         Begin VB.TextBox txtImageLevel 
            Height          =   270
            Left            =   4050
            TabIndex        =   42
            Text            =   "��,��"
            ToolTipText     =   "��������Ӱ�������ĵǼǣ�����ĸ��ȼ�"
            Top             =   990
            Width           =   1035
         End
         Begin VB.Frame fra 
            Caption         =   "¼��ʱ��"
            Height          =   1150
            Index           =   6
            Left            =   5280
            TabIndex        =   38
            Top             =   240
            Width           =   2055
            Begin VB.OptionButton optResultInput 
               Caption         =   "���ǩ����"
               Height          =   240
               Index           =   0
               Left            =   210
               TabIndex        =   41
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optResultInput 
               Caption         =   "���ǩ����"
               Height          =   240
               Index           =   1
               Left            =   210
               TabIndex        =   40
               Top             =   525
               Width           =   1230
            End
            Begin VB.OptionButton optResultInput 
               Caption         =   "�����ӡǰ"
               Height          =   240
               Index           =   2
               Left            =   210
               TabIndex        =   39
               Top             =   810
               Width           =   1290
            End
         End
      End
      Begin VB.Frame fraEditorSetUp 
         Caption         =   "�����ĵ��༭������"
         Height          =   4215
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   7695
         Begin VB.Frame Frame8 
            Caption         =   "�鿴��ʷ����"
            Height          =   1215
            Left            =   240
            TabIndex        =   30
            Top             =   480
            Width           =   7215
            Begin VB.OptionButton optHistoryReportEditor 
               Caption         =   "PACS����༭��"
               Height          =   255
               Index           =   1
               Left            =   4080
               TabIndex        =   32
               Top             =   600
               Width           =   1695
            End
            Begin VB.OptionButton optHistoryReportEditor 
               Caption         =   "���Ӳ����༭��"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   31
               Top             =   600
               Value           =   -1  'True
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����༭��"
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   7730
         Begin VB.OptionButton optReportEditor 
            Caption         =   "PACS���ܱ���༭��"
            Height          =   255
            Index           =   2
            Left            =   4680
            TabIndex        =   28
            Top             =   240
            Width           =   1932
         End
         Begin VB.OptionButton optReportEditor 
            Caption         =   "���Ӳ����༭��"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optReportEditor 
            Caption         =   "PACS����༭��"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   25
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "��������"
         Height          =   3255
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   7695
         Begin VB.Frame Frame7 
            Caption         =   "��ӡ��ʽѡ��ʽ"
            Height          =   1335
            Left            =   4440
            TabIndex        =   33
            Top             =   1800
            Width           =   2895
            Begin VB.CheckBox chkPrintFormat 
               Caption         =   "��ѡ�����ʽ"
               Height          =   255
               Left            =   240
               TabIndex        =   36
               Top             =   960
               Width           =   2295
            End
            Begin VB.OptionButton optPrintFormat 
               Caption         =   "ʼ�ձ���Ĭ�ϸ�ʽ"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   35
               Top             =   600
               Width           =   2415
            End
            Begin VB.OptionButton optPrintFormat 
               Caption         =   "��¼���һ�δ�ӡ��ʽ"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   34
               Top             =   320
               Value           =   -1  'True
               Width           =   2415
            End
         End
         Begin VB.CheckBox chkUntreadPrinted 
            Caption         =   "��˴�ӡ���������"
            Height          =   180
            Left            =   480
            TabIndex        =   27
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox chkSpecialContent 
            Caption         =   "��ʾר�Ʊ������ݣ�"
            Height          =   180
            Left            =   480
            TabIndex        =   23
            Top             =   1080
            Width           =   2055
         End
         Begin VB.ComboBox cboSpecialContent 
            Height          =   300
            Left            =   480
            TabIndex        =   22
            Text            =   "Combo1"
            Top             =   1440
            Width           =   6855
         End
         Begin VB.CheckBox chkExitAfterPrint 
            Caption         =   "��ӡ���˳�"
            Height          =   180
            Left            =   2760
            TabIndex        =   21
            Top             =   720
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Caption         =   "�����ı�������"
            Height          =   1335
            Left            =   480
            TabIndex        =   14
            Top             =   1800
            Width           =   3255
            Begin VB.TextBox txtAdvice 
               Height          =   270
               Left            =   1560
               TabIndex        =   17
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox txtResult 
               Height          =   270
               Left            =   1560
               TabIndex        =   16
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox txtCheckView 
               Height          =   270
               Left            =   1560
               TabIndex        =   15
               Top             =   225
               Width           =   1335
            End
            Begin VB.Label Label3 
               Caption         =   "��    �飺"
               Height          =   255
               Left            =   360
               TabIndex        =   20
               Top             =   975
               Width           =   975
            End
            Begin VB.Label Label2 
               Caption         =   "��������"
               Height          =   255
               Left            =   360
               TabIndex        =   19
               Top             =   615
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "���������"
               Height          =   255
               Left            =   360
               TabIndex        =   18
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.CheckBox chkShowVideoCapture 
            Caption         =   "��ʾ��Ƶ�ɼ�����"
            Height          =   180
            Left            =   2760
            TabIndex        =   13
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtMinImageCount 
            Height          =   270
            Left            =   6840
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "8"
            Top             =   315
            Width           =   495
         End
         Begin VB.CheckBox chkShowImage 
            Caption         =   "��ʾ����ͼ������                               ��������ͼ��ʾ������"
            CausesValidation=   0   'False
            Height          =   180
            Left            =   480
            TabIndex        =   11
            Top             =   360
            Width           =   6375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "����ʾ�˫����"
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Width           =   2415
         Begin VB.OptionButton optWordDblClick 
            Caption         =   "ֱ��д�뱨��"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optWordDblClick 
            Caption         =   "�򿪴ʾ�༭����"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   480
            Width           =   1750
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "����ͼ˫����"
         Height          =   855
         Left            =   2520
         TabIndex        =   4
         Top             =   5640
         Width           =   2895
         Begin VB.OptionButton optImageDblClick 
            Caption         =   "��ͼƬ�༭����"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   6
            Top             =   480
            Width           =   1750
         End
         Begin VB.OptionButton optImageDblClick 
            Caption         =   "ֱ��д�뱨��"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "�ʾ�ģ����ʾ"
         Height          =   855
         Left            =   5400
         TabIndex        =   1
         Top             =   5640
         Width           =   2415
         Begin VB.OptionButton optShowWord 
            Caption         =   "˫������"
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   3
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton optShowWord 
            Caption         =   "ֱ����ʾ"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmReportSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptId As Long   '����ID
Private mblnRefreshed As Boolean

Public Sub zlRefresh(lngDeptID As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngTemp As Long
    Dim lngHintType As Long
    
    mblnRefreshed = True            '���ݱ�ˢ�¹��ˣ����Ա���
    
    mlngDeptId = lngDeptID
    optReportEditor(0).value = True 'Ĭ��ʹ�õ��Ӳ����༭���༭����
    chkShowImage.value = 0          'Ĭ�ϲ���ʾͼ������
    chkShowVideoCapture.value = 0   'Ĭ�ϲ���ʾ��Ƶ�ɼ�����
    
    chkSpecialContent.value = 0     'Ĭ�ϲ���ʾר�Ʊ���
    cboSpecialContent.Enabled = False
    chkExitAfterPrint.value = 0     'Ĭ�ϴ�ӡ���˳�
    optWordDblClick(0).value = True 'Ĭ��˫���ʾ��ֱ��д�뱨��
    optImageDblClick(0).value = True 'Ĭ�ϱ�������ͼ˫����ֱ��д�뱨��
    txtCheckView.Text = "�������"  'Ĭ��Ϊ�������
    txtResult.Text = "������"     'Ĭ��Ϊ������
    txtAdvice.Text = "����"         'Ĭ��Ϊ����
    optShowWord(0).value = True     'Ĭ��Ϊֱ����ʾ�ʾ�ģ��
    chkUntreadPrinted.value = 0     'Ĭ��Ϊ��˴�ӡ���������
    
    chkIgnorePosi.value = 0     '���Խ��������
    chkReportAfterResult.value = 0 '��Ӱ�����Ϊ����
    chkDefaultPosi.value = 0        '��Ͻ��Ĭ������Ϊδ��ѡ
    chkConformDetermine.value = 1       '��������ж�Ĭ��Ϊѡ��
    txtImageLevel.Text = "��,��"     'Ĭ��Ӱ�������ȼ�
    txtReportLevel.Text = "��,��"    'Ĭ�ϱ��������ȼ�
    
    On Error GoTo err
    strSQL = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId)
    
    While Not rsTemp.EOF
        Select Case rsTemp!������
            Case "����༭��"
                If Nvl(rsTemp!����ֵ, 0) = 0 Then
                    optReportEditor(0).value = True
                ElseIf Nvl(rsTemp!����ֵ, 0) = 1 Then
                    optReportEditor(1).value = True
                Else
                    optReportEditor(2).value = True
                End If
            Case "�鿴��ʷ����"
                If Nvl(rsTemp!����ֵ, 0) = 0 Then
                    optHistoryReportEditor(0).value = True
                Else
                    optHistoryReportEditor(1).value = True
                End If
                
            Case "��ʾ����ͼ��"
                chkShowImage.value = Nvl(rsTemp!����ֵ, 0)
            Case "��������ͼ����"
                txtMinImageCount.Text = Nvl(rsTemp!����ֵ, "8")
            Case "��ʾ��Ƶ�ɼ�"
                chkShowVideoCapture.value = Nvl(rsTemp!����ֵ, 0)
            Case "��ӡ���˳�"
                chkExitAfterPrint.value = Nvl(rsTemp!����ֵ, 0)

            Case "��ʾר�Ʊ���"
                chkSpecialContent.value = Nvl(rsTemp!����ֵ, 0)
                cboSpecialContent.Enabled = IIf(chkSpecialContent.value = 1, True, False)
            Case "ר�Ʊ���ҳ"
                cboSpecialContent.Text = Nvl(rsTemp!����ֵ)
            Case "����ʾ�˫������"
                If Nvl(rsTemp!����ֵ, 0) = 0 Then
                    optWordDblClick(0).value = True
                Else
                    optWordDblClick(1).value = True
                End If
            Case "����ͼ˫������"
                If Nvl(rsTemp!����ֵ, 0) = 0 Then
                    optImageDblClick(0).value = True
                Else
                    optImageDblClick(1).value = True
                End If
            Case "�����������"
                txtCheckView.Text = Nvl(rsTemp!����ֵ, "�������")
            Case "����������"
                txtResult.Text = Nvl(rsTemp!����ֵ, "������")
            Case "��������"
                txtAdvice.Text = Nvl(rsTemp!����ֵ, "����")
            Case "��ʾ�ʾ�ʾ��"
                If Nvl(rsTemp!����ֵ, 0) = 0 Then
                    optShowWord(0).value = True
                Else
                    optShowWord(1).value = True
                End If
            Case "��˴�ӡ���������"
                chkUntreadPrinted.value = Nvl(rsTemp!����ֵ, 0)
            Case "��ӡ��ʽѡ��ʽ"
                If Nvl(rsTemp!����ֵ, 0) = 0 Then
                    optPrintFormat(0).value = True
                Else
                    optPrintFormat(1).value = True
                End If
            Case "��ѡ�����ʽ"
                    chkPrintFormat.value = IIf(Nvl(rsTemp!����ֵ, 0), 1, 0)
            Case "��Ͻ����ʾ����"
                lngHintType = Nvl(rsTemp!����ֵ, 0)
                optResultInput(lngHintType).value = True
            Case "��Ͻ��Ĭ������"
                chkDefaultPosi.value = Nvl(rsTemp!����ֵ, 0)
            Case "��Ӱ�����Ϊ����"
                chkReportAfterResult.value = Nvl(rsTemp!����ֵ, 0)
            Case "���Խ��������"
                chkIgnorePosi.value = Nvl(rsTemp!����ֵ, 0)
            Case "��������ж�"
                chkConformDetermine.value = Nvl(rsTemp!����ֵ, 0)
            Case "Ӱ�������ж�"
                chkImageLevel.value = Nvl(rsTemp!����ֵ, 0)
            Case "Ӱ�������ȼ�"
                txtImageLevel.Text = Nvl(rsTemp!����ֵ, "��,��")
                txtImageLevel.Enabled = chkImageLevel.value = 1
            Case "���������ж�"
                chkReportLevel.value = Nvl(rsTemp!����ֵ, 0)
            Case "���������ȼ�"
                txtReportLevel.Text = Nvl(rsTemp!����ֵ, "��,��")
                txtReportLevel.Enabled = chkReportLevel.value = 1
        End Select
        rsTemp.MoveNext
    Wend
    
    If optReportEditor(2).value Then
        fraEditorSetUp.Visible = True
        
    Else
        fraEditorSetUp.Visible = False
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Public Sub zlSave()
    Dim intMatch As Integer
    Dim strSQL As String
    Dim intTxtLen As Integer
    
    On Error GoTo errHand
    
    If mblnRefreshed = False Then Exit Sub          '����û�б�ˢ�£����Բ�����
    
    If txtImageLevel.Enabled Then
        '������״̬�µ� �����滻��Ӣ��״̬
        txtImageLevel.Text = Replace(txtImageLevel.Text, "��", ",")
        
        intTxtLen = Len(txtImageLevel.Text) - Len(Replace(txtImageLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBoxD Me, "Ӱ��ȼ�����Ϊ2�֣����Ϊ4�֣���������д��", vbOKOnly, "��ʾ��Ϣ"
            txtImageLevel.Text = Nvl(GetDeptPara(mlngDeptId, "Ӱ�������ȼ�", "��,��"))
            txtImageLevel.SetFocus
            Exit Sub
        End If
    End If
    
    
    If txtReportLevel.Enabled Then
        '������״̬�µ� �����滻��Ӣ��״̬
        txtReportLevel.Text = Replace(txtReportLevel.Text, "��", ",")
        
        intTxtLen = Len(txtReportLevel.Text) - Len(Replace(txtReportLevel.Text, ",", ""))
        
        If intTxtLen > 3 Or intTxtLen < 1 Then
            MsgBoxD Me, "����ȼ�����Ϊ2�֣����Ϊ4�֣���������д��", vbOKOnly, "��ʾ��Ϣ"
            txtReportLevel.Text = Nvl(GetDeptPara(mlngDeptId, "���������ȼ�", "��,��"))
            txtReportLevel.SetFocus
            Exit Sub
        End If
    End If
    
    If optReportEditor(0).value = True Then         '���Ӳ����༭��
        intMatch = 0
    ElseIf optReportEditor(1).value = True Then     'PACS����༭��
        intMatch = 1
    ElseIf optReportEditor(2).value = True Then     '�����ĵ��༭��
        intMatch = 2
    End If
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '����༭��','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��ʾ����ͼ��','" & chkShowImage.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��������ͼ����','" & txtMinImageCount.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��ʾ��Ƶ�ɼ�','" & chkShowVideoCapture.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��ӡ���˳�','" & chkExitAfterPrint.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��ʾר�Ʊ���','" & chkSpecialContent.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", 'ר�Ʊ���ҳ','" & cboSpecialContent.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If optWordDblClick(0).value = True Then         '����ʾ�˫����ֱ��д�뱨��
        intMatch = 0
    ElseIf optWordDblClick(1).value = True Then     '����ʾ�˫����򿪱༭����
        intMatch = 1
    End If
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '����ʾ�˫������','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If optImageDblClick(0).value = True Then         '����ͼ˫����ֱ��д�뱨��
        intMatch = 0
    ElseIf optImageDblClick(1).value = True Then     '����ͼ˫�����ͼ��༭����
        intMatch = 1
    End If
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '����ͼ˫������','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�����������','" & txtCheckView.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '����������','" & txtResult.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��������','" & txtAdvice.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If optShowWord(0).value = True Then         'ֱ����ʾ�ʾ�ʾ��
        intMatch = 0
    ElseIf optShowWord(1).value = True Then     '˫���������ʾ�ʾ�ʾ��
        intMatch = 1
    End If
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��ʾ�ʾ�ʾ��','" & intMatch & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��˴�ӡ���������','" & chkUntreadPrinted.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If optReportEditor(2) Then
        strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�鿴��ʷ����','" & IIf(optHistoryReportEditor(0).value, 0, 1) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��ӡ��ʽѡ��ʽ','" & IIf(optPrintFormat(0).value, 0, 1) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��ѡ�����ʽ','" & IIf(chkPrintFormat.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��������ж�','" & IIf(chkConformDetermine.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '���Խ��������','" & IIf(chkIgnorePosi.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��Ӱ�����Ϊ����','" & IIf(chkReportAfterResult.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��Ͻ��Ĭ������','" & IIf(chkDefaultPosi.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", 'Ӱ�������ж�','" & IIf(chkImageLevel.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", 'Ӱ�������ȼ�','" & txtImageLevel.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '���������ж�','" & IIf(chkReportLevel.value, 1, 0) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '���������ȼ�','" & txtReportLevel.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��Ͻ����ʾ����','" & IIf(optResultInput(0).value = True, 0, IIf(optResultInput(1).value = True, 1, 2)) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub chkSpecialContent_Click()
    If chkSpecialContent.value = 1 Then
        cboSpecialContent.Enabled = True
    Else
        cboSpecialContent.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    mblnRefreshed = False
    'װ��ר�Ʊ�������
    cboSpecialContent.Clear
    cboSpecialContent.AddItem (Report_Form_frmReportES)
    cboSpecialContent.AddItem (Report_Form_frmReportPathology)
    cboSpecialContent.AddItem (Report_Form_frmReportUS)
    cboSpecialContent.AddItem (Report_Form_frmReportCustom)
End Sub

Private Sub Form_Resize()
    fraReportSetup.Left = (Me.ScaleWidth - fraReportSetup.Width) / 2
End Sub


Private Sub optReportEditor_Click(Index As Integer)
    Dim hService As Long
    Dim hSCManager As Long

On Error GoTo errHandle

    fraEditorSetUp.Visible = Index = 2
    
    Exit Sub
errHandle:
    
End Sub

Private Sub chkImageLevel_Click()
    txtImageLevel.Enabled = chkImageLevel.value = 1
End Sub

Private Sub chkReportAfterResult_Click()
    If chkReportAfterResult.value = vbChecked Then
        chkIgnorePosi.Enabled = False
        chkIgnorePosi.value = vbUnchecked
    Else
        chkIgnorePosi.Enabled = True
    End If
End Sub

Private Sub chkReportLevel_Click()
    txtReportLevel.Enabled = chkReportLevel.value = 1
End Sub
