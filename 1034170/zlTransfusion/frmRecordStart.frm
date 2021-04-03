VERSION 5.00
Begin VB.Form frmRecordStart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ��ʼʱ��"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecordStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2415
      TabIndex        =   3
      Top             =   1245
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   615
      TabIndex        =   2
      Top             =   1245
      Width           =   1100
   End
   Begin VB.TextBox txtStart 
      Height          =   300
      Left            =   1005
      TabIndex        =   1
      Top             =   255
      Width           =   2670
   End
   Begin VB.ComboBox cboOper 
      Height          =   315
      Left            =   1020
      TabIndex        =   0
      Text            =   "cboOper"
      Top             =   615
      Width           =   2670
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Ա"
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   5
      Top             =   660
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼʱ��"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   300
      Width           =   720
   End
End
Attribute VB_Name = "frmRecordStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mOutNurse As OutNurses
Private mstrOper As String
Private mstrDate As String
Private mblnOk As Boolean

Public Function ShowSelect(ByRef ObjOutNurse As OutNurses, ByRef strDate As String, ByRef strOper As String) As Boolean
'������
'  objOutNurse����ʿ
'  strDate����ǰ����ʱ��
'  strOper�������˻�ִ����

    Set mOutNurse = ObjOutNurse
    mstrDate = strDate
    mstrOper = strOper
    mblnOk = False
    
    Me.Show vbModal
    ShowSelect = mblnOk
    If mblnOk Then
        strDate = mstrDate
        strOper = mstrOper
    End If
End Function

Private Sub cboOper_KeyPress(KeyAscii As Integer)
    Dim strTmp As String
    Dim objOutNur As OutNurse
    Dim blnFind As Boolean
     
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        strTmp = cboOper.Text
        If strTmp = "" Then Exit Sub
        If zlCommFun.IsCharChinese(strTmp) Then
            For Each objOutNur In mOutNurse
                If UCase(strTmp) = UCase(objOutNur.����) Then
                    Call cbo.SeekIndex(cboOper, objOutNur.����)
                    blnFind = True
                    Exit For
                End If
            Next
        Else
            For Each objOutNur In mOutNurse
                If UCase(strTmp) = UCase(objOutNur.����) Or UCase(strTmp) = UCase(objOutNur.���) Then
                    strTmp = objOutNur.����
                    Call cbo.SeekIndex(cboOper, strTmp)
                    blnFind = True
                    Exit For
                End If
            Next
        End If
        If Not blnFind Then
            MsgBox "δ�ҵ��ò�����Ա��", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub cmdCancle_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim strDate As String, strOper As String
    Dim i As Long, blnFind As Boolean
    Dim strTmp As String
    
    If Not IsDate(txtStart) Then
        MsgBox "��ʼʱ���ʽ���ԣ�"
        Exit Sub
    End If
    
    strTmp = Trim(cboOper.Text)
    If strTmp = "" Then
         MsgBox "��ѡ�������Ա��", vbInformation, gstrSysName
         Exit Sub
    End If
    
    For i = 0 To cboOper.ListCount - 1
        If strTmp = cboOper.List(i) Then
            blnFind = True
            Exit For
        End If
    Next
    If Not blnFind Then
        MsgBox "δ�ҵ�������Ա:" & strTmp & "��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstrOper = cboOper.List(cboOper.ListIndex)
    mstrDate = Format(CDate(txtStart), "yyyy-MM-dd HH:mm:ss")
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim intIndex As Integer, i As Integer
    Dim objOutNur As OutNurse
    
    Me.cboOper.Clear
    
    intIndex = -1
    cboOper.Clear
    For Each objOutNur In mOutNurse
        Me.cboOper.AddItem objOutNur.����
        If mstrOper = objOutNur.���� And mstrOper <> "" Then
            intIndex = cboOper.NewIndex
        End If
    Next

    If cboOper.ListCount > 0 Then
        If intIndex <= -1 Then
            '�����ʿ�б�û�и�ִ���ˣ�mstrOper������Ĭ��ʹ�õ�ǰ��¼�û�������
            For i = 0 To cboOper.ListCount - 1
                If UserInfo.���� = cboOper.List(i) Then
                    cboOper.ListIndex = i
                    Exit For
                End If
            Next
        Else
            cboOper.ListIndex = intIndex
        End If
    End If
    
    Me.txtStart = mstrDate
    
End Sub
