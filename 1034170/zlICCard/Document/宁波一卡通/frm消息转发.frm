VERSION 5.00
Begin VB.Form frm��Ϣת�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ϣת��"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8835
   Icon            =   "frm��Ϣת��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3810
      Top             =   2640
   End
   Begin VB.CommandButton cmd�˳� 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   405
      Left            =   6870
      TabIndex        =   5
      Top             =   5160
      Width           =   1275
   End
   Begin VB.CommandButton cmd�������� 
      Caption         =   "��������(&S)"
      Height          =   405
      Left            =   5430
      TabIndex        =   4
      Top             =   5160
      Width           =   1275
   End
   Begin VB.TextBox txt�������� 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   4275
      Left            =   1350
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   660
      Width           =   7065
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   2
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl����˵�� 
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1350
      TabIndex        =   1
      Top             =   300
      Width           =   7665
   End
   Begin VB.Label lbl��ǰ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   0
      Top             =   300
      Width           =   720
   End
End
Attribute VB_Name = "frm��Ϣת��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnExit As Boolean
Dim blnProcess As Boolean

Private Sub cmd��������_Click()
    Timer1.Enabled = True
    cmd��������.Enabled = False
End Sub

Private Sub cmd�˳�_Click()
    '����ѵ�ǰ����������󣨱��浽���ݿ��У���������Ӧ״̬������ѯ���Ƿ��˳�
    If blnProcess Then
        If MsgBox("���������Ľ�������,�Ƿ��˳���", vbQuestion + vbYesNo + vbDefaultButton2, "��Ϣת��") = vbNo Then Exit Sub
    End If
    
    blnExit = True
    Unload Me
End Sub

Public Function ���ýӿ�(ByVal str���� As String, ByVal lng���к� As Long, ByVal strFuncName As String, ByVal strURL As String) As Boolean
    Dim strSql As String
    Dim strSoapRequest As String
    Dim objHttp As MSXML2.XMLHTTP
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '���Ϊ���ڴ�����
    Me.lbl����˵��.Caption = "���������ķ������к�=[" & lng���к� & "]���������Ժ�......"
    gstrSQL = "zl_��Ϣ����_Insert('" & str���� & "'," & lng���к� & ",'" & strFuncName & "','" & strURL & "',NULL,1)"
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    
    '��ȡ����������
    strSql = " Select �������� From ��Ϣת�� Where ����='" & str���� & "' And ���к�=" & lng���к� & " Order by �к�"
    Call OpenRecordset(rsTemp, "��ȡ����������", strSql)
    With rsTemp
        Do While Not .EOF
            strSoapRequest = strSoapRequest & IIf(IsNull(!��������), "", !��������)
            .MoveNext
        Loop
    End With
    Me.txt��������.Text = strSoapRequest
    
    '׼������
    Set objHttp = New MSXML2.XMLHTTP
    objHttp.Open "post", strURL, False
    objHttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    objHttp.setRequestHeader "Content-Length", Len(strSoapRequest)
    objHttp.setRequestHeader "SOAPAction", strURL
    
    '���ݷ��ص�״̬��Ϣ���ж��Ƿ�ɹ�
    objHttp.send (strSoapRequest)
    If objHttp.Status <> 200 Then GoTo errHand
    
    '����������д�����ݿ�
    Call WriteData(str����, lng���к�, strFuncName, strURL, objHttp.responseText)
    Me.lbl����˵��.Caption = "���к�=[" & lng���к� & "]�����������"
    
    ���ýӿ� = True
    Exit Function
errHand:
    If Err <> 0 Then
        gstrSQL = "zl_��Ϣ����_Insert('" & str���� & "'," & lng���к� & ",'" & strFuncName & "','" & strURL & "','" & Err.Description & "',3)"
    Else
        gstrSQL = "zl_��Ϣ����_Insert('" & str���� & "'," & lng���к� & ",'" & strFuncName & "','" & strURL & "','������Ϣ��[" & objHttp.Status & "]" & objHttp.responseText & "',3)"
    End If
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
End Function

Private Sub WriteData(ByVal str���� As String, ByVal lng���к� As Long, ByVal strFuncName As String, ByVal strURL As String, ByVal strData As String)
    Dim blnTrans As Boolean
    Dim strRow As String
    Dim intRow As Integer, intCount As Integer
    On Error GoTo errHand
    '����������д�����ݿ�
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    '���±�־
    gcnOracle.Execute "zl_��Ϣ����_Insert('" & str���� & "'," & lng���к� & ",'" & strFuncName & "','" & strURL & "',NULL,9)", , adCmdStoredProc
    
    'д���������
    intCount = Len(strData) \ 1000
    If Len(strData) Mod 1000 <> 0 Then intCount = intCount + 1
    For intRow = 0 To intCount
        strRow = Mid(strData, intRow * 1000 + 1, 1000)
        gcnOracle.Execute "zl_��Ϣ����_Insert('" & str���� & "'," & lng���к� & "," & intRow + 1 & ",'" & Replace(strRow, "'", "''") & "')", , adCmdStoredProc
    Next
    
    gcnOracle.CommitTrans
    blnTrans = False
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    blnExit = False
End Sub

Private Sub Timer1_Timer()
    Dim rsTemp As New ADODB.Recordset
    'ÿ������һ�������ͷ�һ�ο���Ȩ
    
    Timer1.Enabled = False
    gstrSQL = " Select ����,���к�,FUNCNAME,URL From ��Ϣ���� Where nvl(��־,0)=0 Order by ����,���к�"
    Call OpenRecordset(rsTemp, "��ȡ�������嵥")
    With rsTemp
        Do While Not .EOF
            DoEvents
            
            If blnExit Then Exit Sub
            
            blnProcess = True
            Call ���ýӿ�(!����, !���к�, !FuncName, !url)
            blnProcess = False
            
            DoEvents
            .MoveNext
        Loop
    End With
    
    Timer1.Enabled = True
End Sub
