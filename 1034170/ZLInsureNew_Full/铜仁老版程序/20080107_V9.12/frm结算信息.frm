VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm������Ϣ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ϣ"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   ControlBox      =   0   'False
   Icon            =   "frm������Ϣ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshBill 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   4154
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      BackColorSel    =   4194304
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm������Ϣ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

'Modified by ���� 20031218 ���������� ��������
Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select Decode(A.��¼����,1,'��Ԥ��',11,'��Ԥ��',A.���㷽ʽ) ���㷽ʽ,Nvl(A.��Ԥ��,0) ��� " & _
                " From ����Ԥ����¼ A,�����ʻ� B " & _
                " Where A.����ID=B.����ID And B.����=" & gintInsure & " And A.����ID=" & mlng����ID
    Call OpenRecordset(rsTemp, "��ȡ���ν��׽�����Ϣ")
    
    With MshBill
        .Clear
        .Rows = 2
        .Cols = 2
        .TextMatrix(0, 0) = "���㷽ʽ"
        .TextMatrix(0, 1) = "���"
        .ColWidth(0) = 2000
        .ColWidth(1) = 1200
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
    End With
    
    With rsTemp
        Do While Not .EOF
            MshBill.TextMatrix(.AbsolutePosition, 0) = !���㷽ʽ
            MshBill.TextMatrix(.AbsolutePosition, 1) = Format(!���, "#####0.00;-#####0.00; ;")
            If MshBill.Rows - 1 = .AbsolutePosition Then MshBill.Rows = MshBill.Rows + 1
            .MoveNext
        Loop
        If Trim(MshBill.TextMatrix(MshBill.Rows - 1, 0)) = "" Then MshBill.Rows = MshBill.Rows - 1
    End With
End Sub

Public Sub ShowME(Optional ByVal lng����ID As Long = 0)
    mlng����ID = lng����ID
    Me.Show 1
End Sub
