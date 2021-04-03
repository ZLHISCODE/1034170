VERSION 5.00
Begin VB.Form frmAddFile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�����ļ����"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   6525
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2520
      Width           =   6525
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   345
         Left            =   4080
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   345
         Left            =   5280
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblPgs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   45
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.TextBox txtDataFile 
      Height          =   300
      Left            =   1710
      TabIndex        =   2
      Top             =   1560
      Width           =   3945
   End
   Begin VB.CheckBox chkSpaceExtd 
      Caption         =   "�Զ���չ�ռ�"
      Height          =   270
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "AUTOEXTEND ON Next (��ռ��С/10)M"
      Top             =   1965
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.TextBox txtSpaceSize 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "500"
      Top             =   1950
      Width           =   735
   End
   Begin VB.TextBox txtTableSpace 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2160
   End
   Begin VB.TextBox txtFileAmount 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1710
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "4"
      Top             =   1230
      Width           =   300
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAddFile.frx":0000
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   2400
      TabIndex        =   13
      Top             =   1290
      Width           =   2340
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ϊ��ǰ��ռ���������ļ�"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image Img 
      Height          =   480
      Left            =   240
      Picture         =   "frmAddFile.frx":0022
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblDataFile 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��һ���ļ�"
      Height          =   180
      Left            =   720
      TabIndex        =   10
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label lblBakSpace 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݱ�ռ���"
      Height          =   225
      Left            =   480
      TabIndex        =   8
      Top             =   900
      Width           =   1125
   End
   Begin VB.Label lblFileAmount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����     ���ļ�"
      Height          =   195
      Index           =   0
      Left            =   1065
      TabIndex        =   4
      Top             =   1290
      Width           =   1305
   End
   Begin VB.Label lblFileSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ��С                     M"
      Height          =   195
      Left            =   855
      TabIndex        =   9
      Top             =   2010
      Width           =   1785
   End
End
Attribute VB_Name = "frmAddFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ShowAddFile(ByVal strTableSpace As String)
    
    txtTableSpace.Text = strTableSpace
    txtDataFile.Text = GetFileName(, strTableSpace)
    
    Me.Show 1
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function GetFileName(Optional ByVal strFile As String, Optional ByVal strTableSpace As String) As String
    '���ݵ�ǰ�������ļ�����,��ȡ��һ�������ļ�
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    If strFile = "" Then
        strSql = "Select Max(File_Name) Max_File From Dba_Data_Files Where Tablespace_Name =[1]"
        Set rsTmp = OpenSQLRecord(strSql, "��ȡ�����ļ���", strTableSpace)
        strFile = rsTmp!Max_file
    End If
    
    strFile = Left(strFile, InStr(1, strFile, ".DBF") - 1)
    
    If IsNumeric(Right(strFile, 4)) Then
        '����λΪ����,���������� ZLHD2017\2018 ���ְ����Ϊ����ı��������ļ�
        strFile = strFile & "_01.DBF"
    ElseIf IsNumeric(Right(strFile, 3)) Then
        '����Ϊ����
        strTmp = Format(Val(Right(strFile, 3)) + 1, "000")
        strFile = Left(strFile, Len(strFile) - 3) & strTmp & ".DBF"
    ElseIf IsNumeric(Right(strFile, 2)) Then
        '����λΪ����
        strTmp = Format(Val(Right(strFile, 2)) + 1, "00")
        
        strFile = Left(strFile, Len(strFile) - 2) & strTmp & ".DBF"
    ElseIf IsNumeric(Right(strFile, 1)) Then
        '��һλΪ����
        strFile = Left(strFile, Len(strFile) - 1) & Val(Right(strFile, 1)) + 1 & ".DBF"
    Else
        'û������
        strFile = strFile & "01.DBF"
    End If
    
    GetFileName = strFile
End Function


Private Sub cmdOK_Click()
    If AddDatafile(txtTableSpace.Text, txtDataFile.Text, txtFileAmount.Text, txtSpaceSize.Text, chkSpaceExtd.Value) Then
        MsgBox "�����ļ�������ɣ�", , "��ʾ"
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub txtDataFile_GotFocus()
    txtDataFile.SelStart = Len(txtDataFile.Text)
End Sub

Private Sub txtFileAmount_GotFocus()
    txtFileAmount.SelStart = Len(txtFileAmount.Text)
End Sub

Private Sub txtFileAmount_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub

Private Sub txtSpaceSize_GotFocus()
    txtSpaceSize.SelStart = Len(txtSpaceSize.Text)
End Sub

Private Function AddDatafile(ByVal strTableSpace As String, ByVal strFile As String, ByVal intNum As Integer, ByVal lngSize As Long, ByVal blnAutoExtend As Boolean) As Boolean
    'Ϊ��ռ���������ļ�
    '����:strTableSpace - ��ռ�����,strFile - �׸��ļ��� , intNum - ����ļ����� ,lngSize  - ��ʼ�ļ���С, blnAutoExtend - �Ƿ��Զ���չ
    Dim strErrMsg As String, strSql As String
    Dim strNextFile As String, i As Integer, strTmp As String
    
    On Error Resume Next
    
    lblPgs.Caption = "���ڴ��������ļ�������"
    
    For i = 1 To intNum
        If strNextFile = "" Then
            strNextFile = strFile
        Else
            strNextFile = GetFileName(strNextFile)
        End If
        
        lblPgs.Caption = "���ڴ��������ļ�" & strNextFile & "������"
        
        strSql = "Alter TableSpace " & strTableSpace & " Add DataFile '" & strNextFile & "' Size " & lngSize & "M  AutoExtend  " & IIf(blnAutoExtend, "On", "Flase")
        gcnOracle.Execute strSql
        
        If Err.Number <> 0 Then
            strTmp = IIf(InStr(1, strNextFile, "\") > 0, "\", "/")
            strTmp = Mid(strNextFile, InStrRev(strNextFile, strTmp) + 1, InStr(1, strNextFile, ".") - 1)
            strErrMsg = "��������ļ� " & strTmp & "�������� ����ԭ�� ��" & vbNewLine & Err.Description
            
             If MsgBox(strErrMsg & vbNewLine & "�Ƿ�����������������ļ�������ǽ����������ȡ�����˳���ǰ������", vbYesNo, "����") = vbYes Then
                strErrMsg = ""
                Err.Clear
            Else
                lblPgs.Caption = "������ȡ��"
                Exit Function
            End If
        End If
    Next
    
    AddDatafile = True
End Function

Private Sub txtSpaceSize_KeyPress(KeyAscii As Integer)
    OnlyIntCK KeyAscii
End Sub
