VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmQueueStation 
   BorderStyle     =   0  'None
   Caption         =   "�Ŷӽк�"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10575
   ControlBox      =   0   'False
   Icon            =   "frmQueueStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtLocateValue 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7800
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picCallFace 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   5160
      ScaleHeight     =   4455
      ScaleWidth      =   3735
      TabIndex        =   6
      Top             =   600
      Width           =   3735
      Begin XtremeReportControl.ReportControl rptCallList 
         Height          =   3855
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   6800
         _StockProps     =   0
         BorderStyle     =   3
         AllowColumnSort =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption scCallInf 
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "�����б���"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   8421504
      End
   End
   Begin VB.PictureBox picLabel 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10575
      TabIndex        =   0
      Top             =   5670
      Width           =   10575
      Begin VB.CheckBox chkOutQueue 
         Caption         =   "������"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   855
      End
      Begin VB.CheckBox chkOutQueue 
         Caption         =   "����ͣ"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkOutQueue 
         Caption         =   "�����"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
      Begin VB.Label labError 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   0
         Width           =   6975
      End
   End
   Begin VB.Timer tmrBroadCast 
      Interval        =   30000
      Left            =   4440
      Top             =   0
   End
   Begin VB.PictureBox picQueueFace 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   600
      ScaleHeight     =   4455
      ScaleWidth      =   3735
      TabIndex        =   4
      Top             =   600
      Width           =   3735
      Begin XtremeReportControl.ReportControl rptQueueList 
         Height          =   3975
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   7011
         _StockProps     =   0
         BorderStyle     =   3
         AllowColumnSort =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption scQueueInf 
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "�Ŷ��б���"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16761024
         GradientColorDark=   16744576
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   360
      Top             =   720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DkpMain 
      Bindings        =   "frmQueueStation.frx":0CCA
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQueueStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'����˵��
'
'�㲥�����������ݿ��е�������ݣ�ֱ����֯���ݵ����Ŷ��������С��У���ֱ�ӶԶ����е����ݽ��й㲥
'ֱ���������¼���������Ҫ�������ݿ��е����ݣ���ѡ��Զ����е��κ�һ�����ݽ���ֱ��
'˳���������е�����˳�����ֱ���Ĳ���
'�����ִ��˳������ֱ����������˵������ң���ִ�н���
'
'�Է���̨�ɿ���ֱ����˳�����㲥�Ĺ���
'��ҽ��վ�ɿ���ֱ����˳��������Ĺ���
'
'
'ȡ��ԭ�еġ��غ������ܣ��غ��������Ѿ�����˳������ֱ�����Ķ������ݣ�
'
'
'��ɣ�����ǰ�����е���������Ϊ���״̬
'���ţ�������ǰ�����е����ݺ���
'��ͣ����ͣ��ǰ�����е����ݺ���
'�ָ����ָ��Ѿ����š���ͣ����ɵĶ�������
'������½����Ŷӣ���ȷ���������
'
'


Private mcnOracle As ADODB.Connection
Private mstr��������() As String            '�к�ϵͳ����Ҫ������ʾ�����ݶ�������
Private mstrCurrent�������� As String       '��ǰѡ�еĶ�������
Private mIsUnload As Boolean                '�Ƿ��˳�
Private mstrBusinessIds As String           '����������Ѿ����ص�ҵ��ID

Private objVoice As Object                  '�������ж���

Private mlng���з�ʽ As Integer             '���з�ʽ 0 ��ʾ���غ��У�1 ��ʾԶ�˺���
Private mint�����㲥ʱ�䳤�� As Integer     '�������ŵ�ʱ�䳤�ȣ�Ĭ��ֵΪ15��
Private mlng�����㲥���� As Long            '�����㲥�����٣�0-100��,��int��ʱ����Щ�������޷���������
Private mlng�������Ŵ��� As Long            '�����㲥����
Private mstr����վ������ As String          'ִ�к��е�վ������
Private mbln������������ As Boolean         '�Ƿ����ñ��ص��������й���
Private mlngCurPlayCount As Long            '��ǰ�Ѳ��Ŵ���
Private mbln��ʾ�ŶӶ��� As Boolean         '�Ƿ�ʹ����ʾ�豸��ʾ�ŶӶ���
Private mstrShowColumnInf As String         '��ʾ�е�������Ϣ
Private mstrShowCalledColumnInf As String   '������ʾ��������Ϣ
Private mlng���ﲡ������ As Long            '���ﲡ�������Ŷ�
Private mstr�������� As String              '��������
Private mlng��ѯʱ�� As Long                '��ѯʱ�䳤��
Private mlngQueueGroupType As Long          '�Ŷӷ�������
Private mblnIsSelectedCallingList As Boolean '�Ƿ�ѡ���ѽкŶ���
Private mlngQueueFocusRow As Long            '�ŶӶ��н�����
Private mlngCallingFocusRow As Long          '���ж��н�����
Private mstrLocateType As String             '��λ����
Private mobjSquareCard As Object    'һ��ͨ�������㲿��
Private maryKinds() As String
Private mstrPrivs As String                 'Ȩ���ַ���

Private mlngCurrentWorkID As Long           '��ǰѡ�е�ҵ��ID
Private mlngCurrentWorkType As Long         '��ǰҵ������
Private mlngCurrentQueueId As Long          '��ǰ����ID

Private mstrLoginUserName As String
Private mblnFuncState(7) As Boolean         '����״̬ 0-�ָ���1-ֱ��/˳����2-���� ��3-��ͣ��4-��ɾ��5,-�㲥�� 6,���7,����
Private mstr�������� As String
Private mstrҽ������ As String
Private mstrExcludeData As String
Private mintViewDataType As Integer

Private mintDetonatEvent As Integer     '����OnSelectedChange�¼�  0--��ʼֵ�����ã�1--����rptQueueList�Ŷ��б����¼�   2--����rptCallList�����б����¼�
Private mblnNotRefresh As Boolean          'Ϊtrueʱ���Ŷ��б�������б�ִ��ѡ���б任�¼�������ˢ��
'���沼��
Private mlngQueueW1 As Long
Private mlngQueueW2 As Long
Private mlngLEDW As Long
Private mlngLEDH As Long

Private mintIconSize As Integer
Private mblnIsDisplayText As Boolean
Private mblnFirst As Boolean

Public mblnIsShowFindTools As Boolean   '�Ƿ���ʾ���ҹ�����

Private mlngMaxLen As Long '��ȡ�����ŶӺ���ֵ�е����
Private mblnIsGroup As Boolean '�Ŷӽк��б��Ƿ���ʾ����

Private Enum mCol
    �������� = 0: Id: ����ID: �Ŷӱ��: �ŶӺ���:  �Ŷ����: ��������: ����: �������: ���������: ����ID: ����: ҽ������: �Ŷ�״̬: �Ŷ�ʱ��: ����ҽ��: ҵ������: ҵ��ID: ����ʱ��
End Enum

Public Event OnQueueExecuteBefore(ByVal strҵ��ID As String, ByVal byt�������� As Byte, blnCancel As Boolean, strNewQueueName As String)
Public Event OnQueueExecuteAfter(ByVal strҵ��ID As String, ByVal byt�������� As Byte)
Public Event OnRecevieDiagnose(ByVal strҵ��ID As String, ByVal lngҵ������ As Long)
Public Event OnSelectionChanged(ByVal blnIsCallingList As Boolean, objDataRow As XtremeReportControl.ReportRow, cbrMain As XtremeCommandBars.CommandBars)

'Public Sub zlShowMe(cnOracle As ADODB.Connection, str��������() As String, strCurrent�������� As String, lngCurrentWorkID As Long)
'    '���е��±��1��ʼ
'    Call zlRefresh(cnOracle, str��������, strCurrent��������, lngCurrentWorkID)
'
'    Me.Show
'End Sub

''''''''''''��������'''''''''''''''''''''''
Public Sub zlSetToolIcon(ByVal intIconSize As Integer, ByVal blnIsDisplayText As Boolean)
  mintIconSize = intIconSize
  mblnIsDisplayText = blnIsDisplayText
  
  Call Me.cbrMain.Options.SetIconSize(True, mintIconSize, mintIconSize)
  Call Me.cbrMain.RecalcLayout

'  Call SetCommandBarStyle
'  Call InitCommandBars
End Sub



Public Sub zlInitVar(cnOracle As ADODB.Connection, Optional lngSys As Long = 100, _
    Optional intҵ������ As Integer = 0, Optional intValidDays As Integer = 1, _
    Optional strPrivs As String = "", Optional strOption As String = "", Optional blnIsGroup As Boolean = True)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ��ϵͳ����
    '��Σ�strOption-����,�Ժ���չ
    '���ƣ����˺�
    '���ڣ�2010-06-11 11:01:09
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    glngSys = lngSys
    glngModul = 1160
    mstrPrivs = strPrivs
    Set mcnOracle = cnOracle
    mblnIsGroup = blnIsGroup
    
    If Trim(mstrPrivs) = "" Then
        mstrPrivs = GetPrivFunc(glngSys, glngModul)
    End If
    
    Call ClearQueueData(intҵ������, intValidDays)
End Sub


Private Sub ClearQueueData(ByVal intҵ������ As Integer, ByVal intValidDays As Integer)
    Dim strSql As String
    
    On Error GoTo errHandle

    strSql = "ZL_�Ŷ����(" & intҵ������ & "," & intValidDays & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "����Ŷ�����")
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Function GetQueueBusinessDataIDs() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҵ��IDs
    '���:bytType-0-�Һ�;1...
    '����:
    '����:�ɹ�����ҵ��IDs,����ö��ŷ���,��:22,33,44
    '����:���˺�
    '����:2014-03-11 16:48:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    GetQueueBusinessDataIDs = mstrBusinessIds
 

End Function


Private Sub SwitchActiveWindow(ByVal blnIsCalledList As Boolean)
    On Error Resume Next
    
    If blnIsCalledList Then
        scCallInf.GradientColorDark = &HFF8080
        scCallInf.GradientColorLight = &HFFC0C0
        
        scQueueInf.GradientColorDark = &H808080
        scQueueInf.GradientColorLight = &HC0C0C0
    Else
        scQueueInf.GradientColorDark = &HFF8080
        scQueueInf.GradientColorLight = &HFFC0C0
        
        scCallInf.GradientColorDark = &H808080
        scCallInf.GradientColorLight = &HC0C0C0
    End If
End Sub



Private Sub SetReportRecordItem(rriItem As ReportRecord, rsData As ADODB.Recordset)
    Dim i As Integer
    
    On Error GoTo errHandle
    rriItem(mCol.Id).value = rsData("id")
    rriItem(mCol.����ID).value = Nvl(rsData("����ID"))
    
    rriItem(mCol.��������).Caption = rsData("��������") & ":" & IIf(InStr(1, Nvl(rsData("��������")), ":") <= 0, "", Mid(Nvl(rsData("��������")), InStr(1, Nvl(rsData("��������")), ":") + 1))
    rriItem(mCol.��������).value = Nvl(rsData("��������"))

    rriItem(mCol.��������).value = Nvl(rsData("��������"))
    rriItem(mCol.����ID).value = Nvl(rsData("����ID"))
    rriItem(mCol.�Ŷӱ��).value = Nvl(rsData("�Ŷӱ��"))
    rriItem(mCol.�Ŷ����).value = Nvl(rsData("�Ŷ����"))
    rriItem(mCol.�ŶӺ���).value = Lpad(Nvl(rsData("�ŶӺ���")), mlngMaxLen)
    rriItem(mCol.�Ŷ�ʱ��).value = Nvl(rsData("�Ŷ�ʱ��"))
    rriItem(mCol.����ʱ��).value = Nvl(rsData("����ʱ��"))
    rriItem(mCol.�������).value = Nvl(rsData("�������"))
    rriItem(mCol.���������).value = Nvl(rsData("���������"))
    rriItem(mCol.����ҽ��).value = Nvl(rsData("����ҽ��"))
    
    If Nvl(rsData("�������")) = "" Then
        rriItem(mCol.��������).Icon = 807
    Else
        rriItem(mCol.��������).Icon = 3504
    End If
    
    
    If Nvl(rsData("�Ŷ�״̬")) = 1 Then
        rriItem(mCol.�Ŷ�״̬).value = "������"
        For i = 0 To rptQueueList.Columns.Count - 1
            rriItem(i).BackColor = &HC0C0FF
        Next
    ElseIf Nvl(rsData("�Ŷ�״̬")) = 0 Then
        rriItem(mCol.�Ŷ�״̬).value = "�Ŷ���"
        For i = 0 To rptQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbWhite
        Next
    ElseIf Nvl(rsData("�Ŷ�״̬")) = 3 Then
        rriItem(mCol.�Ŷ�״̬).value = "��ͣ"
        For i = 0 To rptQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbYellow
        Next
    ElseIf Nvl(rsData("�Ŷ�״̬")) = 4 Then
        rriItem(mCol.�Ŷ�״̬).value = "���"
        For i = 0 To rptQueueList.Columns.Count - 1
            rriItem(i).BackColor = ColorConstants.vbGreen
        Next
    ElseIf Nvl(rsData("�Ŷ�״̬")) = 7 Then
        rriItem(mCol.�Ŷ�״̬).value = "�Ѻ���"
'        For i = 0 To rptQueueList.Columns.Count - 1
'            rriItem(i).BackColor = &HFFC0C0
'        Next
    Else
        rriItem(mCol.�Ŷ�״̬).value = "������"
        For i = 0 To rptQueueList.Columns.Count - 1
            rriItem(i).BackColor = &HC0C0C0
        Next
    End If
    
    If mlngQueueGroupType = 0 Then
        rriItem(mCol.ҽ������).value = Nvl(rsData("ҽ������"))
    Else
        rriItem(mCol.ҽ������).value = Nvl(rsData("��������")) & ":" & Nvl(rsData("ҽ������"))
    End If

    rriItem(mCol.ҵ������).value = Nvl(rsData("ҵ������"))
    rriItem(mCol.ҵ��ID).value = Nvl(rsData("ҵ��ID"))

    rriItem(mCol.����).value = IIf(Nvl(rsData("����")) = 1, "����", "")
    rriItem(mCol.����).value = Nvl(rsData("����"))
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub
 
 
 Private Sub SetFocusToCalledList()
    On Error Resume Next
    
    'If rptCallList.Visible Then rptCallList.SetFocus
    
    On Error GoTo 0
 End Sub
 
 
 Private Sub SetFocusToQueueList()
    On Error Resume Next
    
    'If rptQueueList.Visible Then rptQueueList.SetFocus
    
    On Error GoTo 0
 End Sub

Public Function zlRefresh(str��������() As String, ByVal strCur�������� As String, ByVal strCurҵ��ID As String, _
    Optional str���� As String = "", Optional strҽ�� As String = "", Optional strExcludeData As String = "", Optional intViewDataType As Integer = 0) As Long
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����ˢ��ָ��ҽ��id�ı������ݣ�����������ṩ�༭����
    '��Σ�str��������():�����ָ����������(��1��ʼ)
    '         strCur��������-��ǰ��������
    '         lngCurҵ��ID-ҵ��ID
    '         str����-����Ϊָ��������,����Ϊ�������:��"һ����,������,..."
    '         strҽ��-����Ϊ�ƶ���ҽ��,���Դ����ҽ��,�ö��ŷָ�,��"����,����,..."
    '         strExcludeData-�ų���ָ��ҵ��ID
    '         intViewDataType������ʾ���ͣ�0��ʾ��ǰ�����µ��������ݣ�
    '                                      1��ʾ����Ϊ��ǰ������ҽ������Ϊ�գ�����ҽ���������ڵ�ǰҽ������������Ϊ�պ�ҽ��Ϊ�յ�����
    '                                      2��ʾ����Ϊ��ǰ���Һ�ҽ������Ϊ�ջ�ҽ���������ڵ�ǰҽ��������
    '                                      3��ʾ��ǰҽ��������
    '���ƣ����˺�
    '���ڣ�2010-06-11 20:54:55
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsLocal As ADODB.Recordset
    Dim rptRecord As ReportRecord
    Dim rptCalling As ReportRecord
    Dim strSql As String, j As Integer, i As Integer, strִ��״̬ As String
    Dim strQueueId As String
    Dim strValues(0 To 10) As String, strValue As String, strUninTable As String
    Dim strFilter As String
        
    err = 0: On Error GoTo errHandle
    
    If mblnNotRefresh Then Exit Function  'ִ���¼�OnSelectionChanged������ˢ��
    
    mstrBusinessIds = ""
    
    mstr�������� = str��������
    'mstrCurrent�������� = strCur�������� 'ˢ��ʱˢ�����еĶ������ݣ�����Ҫ���ǵ�ǰ��������
    'mlngCurrentWorkID = lngCurҵ��ID
    mstrExcludeData = strExcludeData
    mstr�������� = str����
    mstrҽ������ = strҽ��
    mintViewDataType = intViewDataType
    
    If rptQueueList.SelectedRows.Count > 0 Then mlngQueueFocusRow = rptQueueList.SelectedRows(0).Index
    If rptCallList.SelectedRows.Count > 0 Then mlngCallingFocusRow = rptCallList.SelectedRows(0).Index
    
'    Call DebugMsg("ZLQueueInterface", "ZLRefresh", "str����=" & str���� & ", strҽ��=" & strҽ�� & ", strExcludeData=" & strExcludeData & ", intViewDataType=" & intViewDataType)
    
    strFilter = "": strValue = "": j = 0: strUninTable = ""
    If SafeArrayGetDim(mstr��������) > 0 Then
        For i = 1 To UBound(mstr��������)
            If Trim(mstr��������(i)) <> "" Then
                If j > 10 Then
                    strFilter = strFilter & " Or A.�������� ='" & str��������(i) & "'"
                Else
                    If zlCommFun.ActualLen(strValue) > 2000 Then
                         strValues(j) = Mid(strValue, 2)
                         strUninTable = strUninTable & " Union ALL  Select  Column_Value as �������� From Table(Cast(f_Str2list([" & j + 4 & "]) As zlTools.t_Strlist))  " & vbCrLf
                         strValue = "": j = j + 1
                    End If
                    strValue = strValue & "," & str��������(i)
                End If
            End If
        Next i
        If strValue <> "" Then
            strValues(j) = Mid(strValue, 2)
            strUninTable = strUninTable & " Union ALL  Select  Column_Value as �������� From Table(Cast(f_Str2list([" & j + 4 & "]) As zlTools.t_Strlist))  " & vbCrLf
        End If
    End If
    
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
       'strUninTable = " Select  Column_Value as �������� From Table(Cast(f_Str2list([3]) As zlTools.t_Strlist)) "
       '���û�д��ݶ������ƣ���ֱ���˳�
'       MsgBox "û�п���ʾ�ĽкŶ�����Ϣ����������Ŷӿ������á�", vbInformation, "�Ŷӽк�ϵͳ"
        labError.Caption = "û�п���ʾ�ĽкŶ�����Ϣ����������Ŷӿ�������"
        Exit Function
    End If
    
    If strFilter <> "" Then strFilter = "( " & Mid(strFilter, 4) & ")"
    
    strִ��״̬ = ","
    
    If chkOutQueue(0).value = vbUnchecked Then strִ��״̬ = strִ��״̬ & "," & "2"
    If chkOutQueue(1).value = vbUnchecked Then strִ��״̬ = strִ��״̬ & "," & "3"
    If chkOutQueue(2).value = vbUnchecked Then strִ��״̬ = strִ��״̬ & "," & "4"
    
'    Call DebugMsg("ZLQueueInterface", "ZLRefresh", "strִ��״̬=" & strִ��״̬)
    
   
    'Ϊ��֧�ָ��ƣ���Ҫ��number���͵��ֶν���ת��������ʹ��to_Number��ʽ
    strSql = "" & _
    "   Select /*+ Rule*/  to_Number(A.ID) as ID, to_Number(a.����id) as ����id, A.��������, A.�Ŷ����, to_Number(A.ҵ������) as ҵ������, to_Number(A.ҵ��ID) as ҵ��ID," & _
    "           to_Number(����ID) as ����ID, x.���� as ��������, �ŶӺ���, �Ŷӱ��,��������,����,ҽ������," & _
    "            (select ���� from ��Ա�� J, �ϻ���Ա�� K where J.ID=K.��ԱID and K.�û���=A.����ҽ��) as ����ҽ��, " & _
    "           to_Number(����) as ����, to_Number(�������) as �������, To_Char(�Ŷ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') as �Ŷ�ʱ��, To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') as ����ʱ��,to_Number(�Ŷ�״̬) as �Ŷ�״̬, " & _
                IIf(mlng���ﲡ������ = 1, "to_number(nvl(�������, 9999999999)) as ���������", "0 as ���������") & _
    "   From �ŶӽкŶ��� a, ���ű� x " & IIf(strUninTable <> "", ", (" & strUninTable & ") b ", "") & _
                IIf(mintViewDataType = 1, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(mintViewDataType = 2, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(mintViewDataType = 3, ", Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D", "") & _
    "   Where   (nvl(�Ƿ��ʱ��, 0)=0 and A.�Ŷ�ʱ�� <= trunc(sysdate + 1) - 1/24/60/60 or nvl(�Ƿ��ʱ��, 0)=1 and sysdate > a.�Ŷ�ʱ��) " & IIf(strUninTable <> "", " and a.��������=b.�������� ", "") & " and instr([3],A.�Ŷ�״̬)=0  and x.ID=a.����ID  " & _
                IIf(mintViewDataType = 1, " and  ((a.����=C.Column_Value and a.ҽ������ is null) or a.ҽ������=D.Column_Value or (a.���� is null and a.ҽ������ is null))", "") & _
                IIf(mintViewDataType = 2, " and (a.����=C.Column_Value and (a.ҽ������ is Null or a.ҽ������=D.Column_Value)) ", "") & _
                IIf(mintViewDataType = 3, " and a.ҽ������=D.Column_Value", "") & _
    "           " & strFilter & _
    "   Order by ���� Desc, �Ŷ�״̬ desc, �Ŷ����, ���������, �Ŷ�ʱ��, �ŶӺ��� "
    

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ѯ����", mstr��������, mstrҽ������, strִ��״̬, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    Set rsLocal = zlDatabase.CopyNewRec(rsTemp)
    
    'ɾ����Ҫ�ų�������,����ȡʵ���ŶӺ���ֵ�����
    While Not rsLocal.EOF
        If InStr(1, strExcludeData, rsLocal!ҵ������ & ":" & rsLocal!ҵ��ID) > 0 Then
            rsLocal.Delete
        End If
        
        If LenB(StrConv(Trim(Nvl(rsLocal("�ŶӺ���"))), vbFromUnicode)) > mlngMaxLen Then
            mlngMaxLen = LenB(StrConv(Trim(Nvl(rsLocal("�ŶӺ���"))), vbFromUnicode))
        End If
        
        rsLocal.MoveNext
    Wend
    
    rsLocal.Sort = "��������, �Ŷ�״̬ desc,���� Desc, �Ŷ����, ���������, �Ŷ�ʱ��, �ŶӺ���"
    If rsLocal.RecordCount > 0 Then rsLocal.MoveFirst
    
    Call rptQueueList.Records.DeleteAll
    Call rptCallList.Records.DeleteAll
        
    While Not rsLocal.EOF

        If rsLocal("�Ŷ�״̬") = 7 Or rsLocal("�Ŷ�״̬") = 1 Then
            Set rptCalling = rptCallList.Records.Add
            For j = 0 To Me.rptCallList.Columns.Count - 1
                rptCalling.AddItem ""
            Next
            
            Call SetReportRecordItem(rptCalling, rsLocal)
        Else
            Set rptRecord = rptQueueList.Records.Add
            For j = 0 To Me.rptQueueList.Columns.Count - 1
                rptRecord.AddItem ""
            Next
            
            Call SetReportRecordItem(rptRecord, rsLocal)
        End If
        
        If mstrBusinessIds <> "" Then mstrBusinessIds = mstrBusinessIds & ","
        mstrBusinessIds = mstrBusinessIds & Nvl(rsLocal!ҵ��ID)
        
        rsLocal.MoveNext
    Wend
    
    rptQueueList.Populate
    rptCallList.Populate
        
        
    On Error GoTo errContinue
    
    '�ָ�ѡ����Ŷ�����
    If mlngQueueFocusRow >= rptQueueList.Rows.Count Then
        mlngQueueFocusRow = IIf(rptQueueList.Rows.Count <= 0, -1, rptQueueList.Rows.Count - 1)
    End If
    
    If mlngQueueFocusRow > -1 Then
        rptQueueList.Rows(mlngQueueFocusRow).Selected = True
    End If
    
    '�ָ�ѡ��ĺ�������
    If mlngCallingFocusRow >= rptCallList.Rows.Count Then
        mlngCallingFocusRow = IIf(rptCallList.Rows.Count <= 0, -1, rptCallList.Rows.Count - 1)
    End If
        
    If mlngCallingFocusRow > -1 Then
        rptCallList.Rows(mlngCallingFocusRow).Selected = True
    End If
        

    '�ָ������б�
    If mblnIsSelectedCallingList Then
        Call SetFocusToCalledList
        
        Call SwitchActiveWindow(mblnIsSelectedCallingList)
    Else
        Call SetFocusToQueueList
        
        Call SwitchActiveWindow(mblnIsSelectedCallingList)
    End If
    
errContinue:
    
    '��ʾ�ŶӶ���
    Call ShowQueue
    
    zlRefresh = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    
    Call SaveErrLog
    
    zlRefresh = err.Number
End Function


Public Sub zlCommandBarSet(ByVal intFuncType As Integer, ByVal blnUseState As Boolean)
'************************************************************************************
'
'���ù���״̬
'
'intFuncType���������� 0-�ָ���1-ֱ��/˳����2-���� ��3-��ͣ��4-��ɾ��5,-�㲥, 6,-����, 7-����
'blnUseState���Ƿ�����
'
'************************************************************************************
    If (intFuncType >= 0) And (intFuncType <= 7) Then
        mblnFuncState(intFuncType) = blnUseState
    End If
End Sub


Public Function zlQueueExec(str��ǰ������ As String, lngҵ������ As Long, strҵ��ID As String, byt�������� As Byte) As Boolean
'*************************************************************************************
'
'ִ�нк���ز���
'
'str��ǰ����������Ҫ�����Ķ�������,������ȷ�����һ���ҽ����ʱ��ʹ�ÿ���ID��Ϊ��������
'
'lngҵ��ID����ʾ��ǰҵ���ID����
'
'byt�������ͣ��кŲ��������� 0-�ָ���1-ֱ��/˳����Lngҵ��ID=0Ϊ˳������2-���� ��3-��ͣ��4-��ɾ��5,-�㲥 6,-����
'
'*************************************************************************************

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngQueueId As Long
    Dim blnFind As Boolean
    Dim i As Integer
            
    On Error GoTo errHandle
    
    zlQueueExec = False
    mstrCurrent�������� = str��ǰ������
        
    Select Case byt��������
        Case 0, 1, 2, 3, 4, 6
            strSql = "ZL_�ŶӽкŶ���_����('" & str��ǰ������ & "'," & byt�������� & ",'" & mstrLoginUserName & "','" & mstr����վ������ & "'," & lngҵ������ & ",'" & strҵ��ID & "')"
            zlDatabase.ExecuteProcedure strSql, "�Ŷӽк�"
        Case 5

            
            strSql = "select ID from �ŶӽкŶ��� where ��������=[1] and ҵ��ID=[2] and ҵ������=[3]"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�㲥", str��ǰ������, strҵ��ID, lngҵ������)
            
            While Not rsTemp.EOF
                lngQueueId = rsTemp!Id
        
                strSql = "ZL_�Ŷ���������_INSERT(" & lngQueueId & ",'" & mstr����վ������ & "', 1)"
                Call zlDatabase.ExecuteProcedure(strSql, "�㲥")
                
                rsTemp.MoveNext
            Wend
    End Select

        
    '����б��д��ڸ����ݣ���λ������
    blnFind = False
    For i = 0 To rptQueueList.Rows.Count - 1
        If rptQueueList.Rows(i).GroupRow <> True Then
            If rptQueueList.Rows(i).Record(mCol.ҵ��ID).value = strҵ��ID _
                And rptQueueList.Rows(i).Record(mCol.ҵ������).value = lngҵ������ _
                And rptQueueList.Rows(i).Record(mCol.��������).value = mstrCurrent�������� Then
            
                rptQueueList.Rows(i).Selected = True
                blnFind = True
            
                Exit For
            End If
        End If
    Next i
    
    '���Ѻ����б��в�������
    If Not blnFind Then
        Call SetFocusToCalledList
        
        For i = 0 To rptCallList.Rows.Count - 1
            If rptCallList.Rows(i).GroupRow <> True Then
                If rptCallList.Rows(i).Record(mCol.ҵ��ID).value = strҵ��ID _
                    And rptCallList.Rows(i).Record(mCol.ҵ������).value = lngҵ������ _
                    And rptCallList.Rows(i).Record(mCol.��������).value = mstrCurrent�������� Then
                
                    rptCallList.Rows(i).Selected = True
                
                    Exit For
                End If
            End If
        Next i
    End If
    
    
    zlQueueExec = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function CheckQueueDataIsHas(ByVal lngQueueId As Long) As Boolean
'***********************************************************************
'�����������Ƿ����
'
'������
'lngQueueId����Ҫ���м��Ķ���ID
'***********************************************************************

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '�ж϶���ID�Ƿ��Ѿ�����
    strSql = "select /*+ RULE*/ ID from �ŶӽкŶ��� where Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ѯ�к������Ƿ����", mlngCurrentQueueId)
    
    CheckQueueDataIsHas = Not rsTemp.EOF
    
    Exit Function
errHandle:
    CheckQueueDataIsHas = False
    If ErrCenter = 1 Then Resume
End Function


Private Function CheckIsSelectedData() As Boolean
    On Error Resume Next
    
    'ȡ�Ŷ��б�����
    If mblnIsSelectedCallingList = False Then
        If rptQueueList.SelectedRows.Count = 0 Then
            If rptQueueList.Rows.Count > 0 Then
                rptQueueList.Rows(1).Selected = True
                
                 Call rptQueueList_SelectionChanged
            Else
                CheckIsSelectedData = False
                Exit Function
            End If
        Else
            'ѡ�е��в����Ƿ�����,����Ƿ����У���Ҫ���õ��÷����µĵ�һ��
            If rptQueueList.SelectedRows(0).GroupRow = True Then
                If rptQueueList.SelectedRows(0).Childs.Count > 0 Then
                    rptQueueList.SelectedRows(0).Childs(0).Selected = True
                    
                    Call rptQueueList_SelectionChanged
                Else
                    CheckIsSelectedData = False
                    Exit Function
                End If
            Else
                Call rptQueueList_SelectionChanged
            End If
        End If
    Else
    'ȡ�Ѻ����б�����
        If rptCallList.SelectedRows.Count = 0 Then
            If rptCallList.Rows.Count > 0 Then
                rptCallList.Rows(1).Selected = True
                
                Call rptQueueList_SelectionChanged
            Else
                CheckIsSelectedData = False
                Exit Function
            End If
        Else
            'ѡ�е��в����Ƿ�����,����Ƿ����У���Ҫ���õ��÷����µĵ�һ��
            If rptCallList.SelectedRows(0).GroupRow = True Then
                If rptCallList.SelectedRows(0).Childs.Count > 0 Then
                    rptCallList.SelectedRows(0).Childs(0).Selected = True
                    
                    Call rptCallList_SelectionChanged
                Else
                    CheckIsSelectedData = False
                    Exit Function
                End If
            Else
                Call rptCallList_SelectionChanged
            End If
        End If
    End If
    
    CheckIsSelectedData = True
End Function


Private Sub Menu_View_Locate_Type_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    
    mstrLocateType = Split(Control.Caption, "(")(0)
    cbrMain.RecalcLayout

    txtLocateValue.SetFocus
    txtLocateValue.Text = ""
    txtLocateValue.PasswordChar = ""
End Sub


Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnIsHasData As Boolean
    Dim i As Integer
    
    'ִ�й���������
    On Error GoTo errHand
    
    labError.Caption = ""
    
    Select Case Control.Id
        Case conMenu_Queue_CallThis, conMenu_Queue_CallNext, _
             conMenu_Queue_CallFirst, conMenu_Queue_Restore, _
             conMenu_Queue_ReCall, conMenu_Queue_Abandon, _
             conMenu_Queue_Refresh, conMenu_Queue_Setup, _
             conMenu_Queue_Update, conMenu_Queue_Broadcast, _
             conMenu_Queue_Pause, conMenu_Queue_Finaled, _
             conMenu_Queue_Find, conMenu_Queue_ComeBack, _
             conMenu_Queue_RecDiagnose
             
        Case conMenu_Queue_LocateType * 10# + 1 To conMenu_Queue_LocateType * 10# + 50   '��λ
            Call Menu_View_Locate_Type_click(Control)
            Exit Sub
        Case Else
            Exit Sub
    End Select
    
    
    '���Ϊ˳�������������ʱ����ֱ�ӷ���rptQueueList�б�(˳����ˢ�£������ö�����Ҫ�����б�)
    If Control.Id <> conMenu_Queue_CallNext _
        And Control.Id <> conMenu_Queue_Refresh _
        And Control.Id <> conMenu_Queue_Setup _
        And Control.Id <> conMenu_Queue_Find Then
    
        If Not CheckIsSelectedData Then
            MsgBox "û��ѡ����Ҫִ�е����ݣ�����ִ�иò�����", vbInformation, "�Ŷӽк�ϵͳ"
            Exit Sub
        End If
        
        If Not CheckQueueDataIsHas(mlngCurrentQueueId) Then
            MsgBox "���ݲ����ڻ��ѱ�ִ�У�������ˢ�²�����", vbInformation, "�Ŷӽк�ϵͳ"
            Exit Sub
        End If
    End If
        
    
    Select Case Control.Id
        Case conMenu_Queue_CallThis 'ֱ��
            '---
            Call comMenu_ֱ��
            
        Case conMenu_Queue_RecDiagnose '����
            '---
            Call comMenu_����
            
        Case conMenu_Queue_Broadcast '�㲥  ���ִ�С��㲥��������Ҫ�����ݽ���ˢ�²���
            '---
            Call comMenu_�㲥
            
            Exit Sub
        Case conMenu_Queue_CallFirst    '����
            '---
            Call comMenu_����
        
        Case conMenu_Queue_Restore    '�ָ�
            '---
            Call comMenu_�ָ�
            
        Case conMenu_Queue_Abandon      '����
            '---
            Call comMenu_����
            
        Case conMenu_Queue_Pause       '��ͣ
            '---
            Call comMenu_��ͣ
            
        Case conMenu_Queue_Finaled      '���
            '---
            Call comMenu_���
                        
'        Case conMenu_Queue_Refresh      'ˢ�� �ô�����Ҫ����ˢ�£���ִ���κβ����󣬻��ڸù��̵�������ˢ��
'            Call comMenu_ˢ��

        Case conMenu_Queue_Find     '����
            Call comMenu_����
            
        Case conMenu_Queue_CallNext '˳��
            Call comMenu_˳��
        
        Case conMenu_Queue_Update       '�޸�
            Call comMenu_�޸�
            
        Case conMenu_Queue_Setup        '����  ����ǡ����á�����������Ҫ�����ݽ���ˢ��
            Call comMenu_����
            
            Exit Sub
    End Select
    
    Call zlRefresh(mstr��������, mstrCurrent��������, mlngCurrentWorkID, mstr��������, mstrҽ������, mstrExcludeData, mintViewDataType)
    
    
    '��ִ��˳������ֱ��֮����Ҫ���������õ������б�
    If Control.Id = conMenu_Queue_CallThis Or Control.Id = conMenu_Queue_CallNext Then
        For i = 0 To rptCallList.Rows.Count - 1
            If rptCallList.Rows(i).GroupRow <> True Then
                If rptCallList.Rows(i).Record(mCol.ҵ��ID).value = mlngCurrentWorkID And rptCallList.Rows(i).Record(mCol.ҵ������).value = mlngCurrentWorkType Then
                    rptCallList.Rows(i).Selected = True
                    
                    Call rptCallList_SelectionChanged
                    Call SetFocusToCalledList
                    
                    mblnIsSelectedCallingList = True
                    
                    Call SwitchActiveWindow(mblnIsSelectedCallingList)

                    Exit For
                End If
            End If
        Next i
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Sub zlDefCommandBars(ByVal cbsThis As Object)
   '������������ť
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    On Error GoTo errHandle
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons

    '�༭�˵�:���ڹ����˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�Ŷ�(&C)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.Id = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        If InStr(mstrPrivs, "ֱ��") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallThis, "ֱ��"): cbrControl.IconId = 732
        End If
        
        If InStr(mstrPrivs, "˳��") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallNext, "˳��"): cbrControl.IconId = 744: cbrControl.ToolTipText = "��˳�������һ��"
        End If
        
'        If InStr(mstrPrivs, "�غ�") > 0 Then
'            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_ReCall, "�غ�"): cbrControl.IconId = 3014: cbrControl.ToolTipText = "�ٴκ���"
'        End If
        
        If InStr(mstrPrivs, "�㲥") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Broadcast, "�غ�"): cbrControl.IconId = 745
        End If
    End With

    '����������
    '-----------------------------------------------------
    Set cbrToolBar = cbsThis(2)
    For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
        If Val(Left(cbrControl.Id, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.Id, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    
    With cbrToolBar.Controls
        If InStr(mstrPrivs, "ֱ��") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallThis, "ֱ��", cbrControl.Index + 1): cbrControl.IconId = 732: cbrControl.ToolTipText = "ֱ�Ӻ��е�ǰ����"
        
            cbrControl.BeginGroup = True
        End If
        
        If InStr(mstrPrivs, "˳��") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallNext, "˳��", cbrControl.Index + 1): cbrControl.IconId = 744: cbrControl.ToolTipText = "��˳�������һ��"
        End If
        
'        If InStr(mstrPrivs, "�غ�") > 0 Then
'            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_ReCall, "�غ�", cbrControl.Index + 1): cbrControl.IconId = 3014: cbrControl.ToolTipText = "�ٴκ���"
'
'            cbrControl.BeginGroup = True
'        End If
        
        If InStr(mstrPrivs, "�㲥") > 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Broadcast, "�غ�", cbrControl.Index + 1): cbrControl.IconId = 745
        End If
    End With

    '����Ŀ����
    '-----------------------------------------------------
    With cbsThis.KeyBindings

    End With

    '���ò���������
    '-----------------------------------------------------
    With cbsThis.Options

    End With
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub
'''''''''''''''''''''''''''''''''''''''''''''

Private Sub SetCommandBarStyle()
    On Error Resume Next
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, mintIconSize, mintIconSize
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    'Me.cbrMain.ActiveMenuBar.Visible = False
End Sub

Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrMenuBar As CommandBarPopup
    
    On Error GoTo errHandle
    '-----------------------------------------------------
    
    '�ŶӺ��й���������
    Call cbrMain.DeleteAll
    Set cbrToolBar = Me.cbrMain.Add("���й�����", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    
        
    Call cbrToolBar.Controls.DeleteAll
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Refresh, "ˢ��"): cbrControl.IconId = 791
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Broadcast, "�غ�"): cbrControl.IconId = 745
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallThis, "ֱ��"): cbrControl.IconId = 732
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallNext, "˳��"): cbrControl.IconId = 744: cbrControl.ToolTipText = "��˳�������һ��"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
       
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_RecDiagnose, "����"): cbrControl.IconId = 8264: cbrControl.ToolTipText = "�Ա������˽��н��ﴦ��"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallFirst, "����"): cbrControl.IconId = 216: cbrControl.ToolTipText = "����Ϊ���Ⱥ���"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Restore, "�ָ�"): cbrControl.IconId = 252: cbrControl.ToolTipText = "�����ݻָ����Ŷ�״̬"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Pause, "��ͣ"): cbrControl.IconId = 746
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Abandon, "����"): cbrControl.IconId = 8113: cbrControl.ToolTipText = "��������"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Finaled, "���"): cbrControl.IconId = 747
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Find, "����"): cbrControl.IconId = 721
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
                       
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Update, "�޸�"): cbrControl.IconId = 3003: cbrControl.ToolTipText = "�޸��Ŷ���Ϣ"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Setup, "����"): cbrControl.IconId = 181: cbrControl.ToolTipText = "��������"
        cbrControl.Caption = IIf(mblnIsDisplayText, cbrControl.Caption, "")
        
        
        
        '�����
        '����
        '���￨��
        'ҽ����
        
        Set cbrMenuBar = .Add(xtpControlPopup, conMenu_Queue_LocateType, "�ŶӺ�")
            cbrMenuBar.Id = conMenu_Queue_LocateType
            cbrMenuBar.Flags = xtpFlagRightAlign
            cbrMenuBar.Category = "CallFind"
        
        
        
        Set cbrCustom = .Add(xtpControlCustom, conMenu_Queue_LocateValue, "��λ����")
            cbrCustom.Handle = txtLocateValue.hwnd
            cbrCustom.Flags = xtpFlagRightAlign
            cbrCustom.Style = xtpButtonIconAndCaption
            cbrCustom.Category = "CallFind"
        
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        If (cbrControl.Type = xtpControlButton) Or (cbrControl.Type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
        If cbrControl.Category = "" Then cbrControl.Category = "Main" '���ó�������˵�
    Next
    
    cbrToolBar.Position = xtpBarTop
    
    
    

    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitFaceScheme()
    '��ʼ���沼��
    
    Dim Pane1 As Pane, Pane2 As Pane, pane3 As Pane
    
    On Error GoTo errHandle
    
    With Me.DkpMain
        .CloseAll
        .SetCommandBars cbrMain
        
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = DkpMain.CreatePane(0, IIf(mlngQueueW1 < 1000, 1000, mlngQueueW1), _
                Me.Height, _
                DockLeftOf, Nothing)
                
    Pane1.Title = "�Ŷ��б�"
    Pane1.Tag = 0
    Pane1.Handle = picQueueFace.hwnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    
    Set Pane2 = DkpMain.CreatePane(1, IIf(mlngQueueW2 < 1000, 1000, mlngQueueW2), _
                Me.Height, _
                DockRightOf, Nothing)
         
    
    Pane2.Title = "�����б�"
    Pane2.Tag = 1
    Pane2.Handle = picCallFace.hwnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    
    
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlExecuteCommandBars(Control)
End Sub


Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    Dim aryKindInfo() As String
    
On Error Resume Next
    If CommandBar.Parent Is Nothing Then Exit Sub
    

    Select Case CommandBar.Parent.Id
        Case conMenu_Queue_LocateType
            With CommandBar.Controls
                If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                    
                    For i = 0 To UBound(maryKinds)
                        aryKindInfo = Split(maryKinds(i), "|")
                        If Trim(aryKindInfo(1)) <> "" Then
                            Set objControl = .Add(xtpControlButton, conMenu_Queue_LocateType * 10# + i + 1, aryKindInfo(1) & "(&" & IIf(i >= 9, Chr(65 + i - 9), i + 1) & ")"): objControl.Category = "CallFind"
                            If i = 0 Then objControl.Checked = True
                        End If
                    Next i
                End If
            End With
    End Select
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlUpdateCommandBars(Control)
End Sub

Private Sub chkOutQueue_Click(Index As Integer)
    Call comMenu_ˢ��
End Sub

Private Sub Form_Activate()
        '��ʾ�ŶӶ���
        If mblnFirst Then
        Call ShowQueue
                mblnFirst = False
        End If
End Sub


Private Sub Form_Load()
    Dim strKinds As String
    
    '��ǰ��½���û���
    mstrLoginUserName = GetUserName()
    mblnFirst = True
    mintIconSize = 24
    mblnIsDisplayText = True
    mIsUnload = False
    mblnIsSelectedCallingList = False
    mlngQueueFocusRow = -1
    mlngCallingFocusRow = -1
    
    mintDetonatEvent = 0
    mblnNotRefresh = False
    
    '���������㲿��
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")

    '��ʼ�������㲿��
    mobjSquareCard.zlInitComponents Me, G_LNG_QUEUEMANAGE_MODULENUM, glngSys, mstrLoginUserName, mcnOracle

    strKinds = ""
    strKinds = strKinds & "���￨|���￨|-1;"
    strKinds = strKinds & "����|����|-1;"
    strKinds = strKinds & "�����|�����|-1;"
    strKinds = strKinds & "ҽ����|ҽ����|-1;"
    strKinds = strKinds & "�ŶӺ�|�ŶӺ�|-1;"

    '��ȡ�ſ������Ϣ
    maryKinds = Split(mobjSquareCard.zlGetIDKindStr(strKinds), ";")
    
    
    Set objVoice = Nothing

    Call InitLocalParas(False)
    Call SetCommandBarStyle
    Call InitCommandBars
    Call InitFaceScheme
    Call InitQueueList
    
    '���Կؼ�λ��
    Call picLabel_Resize
    
End Sub

Private Function GetUserName() As String
'************************************************
'
'ȡ�õ�ǰ��½���û���
'
'************************************************
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
        Set rsTmp = zlDatabase.GetUserInfo
        
        If Not rsTmp.EOF Then
            GetUserName = IIf(IsNull(rsTmp!�û���), "", rsTmp!�û���)
        End If
        
    Exit Function
errHandle:
    GetUserName = ""
    If ErrCenter = 1 Then Resume
End Function


Private Sub InitLocalParas(blnIsLISForm As Boolean)
    Dim strReg As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    strReg = "����ȫ��\�Ŷӽк�"
        
    
    mint�����㲥ʱ�䳤�� = Val(zlDatabase.GetPara("�����㲥ʱ�䳤��", glngSys, glngModul, "15"))
    mlng�����㲥���� = Val(zlDatabase.GetPara("�����㲥����", glngSys, glngModul, "60"))
    mlng�������Ŵ��� = Val(zlDatabase.GetPara("�������Ŵ���", glngSys, glngModul, "1"))
    mbln������������ = zlDatabase.GetPara("������������", glngSys, glngModul, "1")
    mstrShowColumnInf = zlDatabase.GetPara("������ʾ��", glngSys, glngModul, "����,��������,�Ŷ�״̬")
    mstrShowColumnInf = Replace(mstrShowColumnInf, "��", ",")
    mstrShowColumnInf = "," & mstrShowColumnInf & ","
    
    mstrShowCalledColumnInf = zlDatabase.GetPara("����������ʾ��", glngSys, glngModul, "����,��������")
    mstrShowCalledColumnInf = Replace(mstrShowCalledColumnInf, "��", ",")
    mstrShowCalledColumnInf = "," & mstrShowCalledColumnInf & ","
               
    
    mlng���з�ʽ = Val(zlDatabase.GetPara("�кŷ�ʽ", glngSys, glngModul, "0"))
    
    If mlng���з�ʽ Then
        mstr����վ������ = zlDatabase.GetPara("Զ�˺���վ��", glngSys, glngModul, "")
        
        '���Ϊ�վͱ�ʾΪ����վ��
        If Trim(mstr����վ������) = "" Then
          mstr����վ������ = AnalyseComputer
        End If
    Else
        mstr����վ������ = AnalyseComputer
    End If
    
    
    mbln��ʾ�ŶӶ��� = zlDatabase.GetPara("��ʾ�ŶӶ���", glngSys, glngModul, "1")
    plngLEDModal = zlDatabase.GetPara("��ʾ�豸���", glngSys, glngModul, "101")
    
    

    mstrLocateType = GetSetting("ZLSOFT", strReg, "��λ��ʽ", "���￨��")
    
    mlng���ﲡ������ = zlDatabase.GetPara("���ﲡ���Ƿ�����", glngSys, glngModul, "1")
    mlngQueueGroupType = zlDatabase.GetPara("�Ŷӷ�������", glngSys, glngModul, "0")
    
    
    mstr�������� = zlDatabase.GetPara("��������", glngSys, glngModul, "ϵͳĬ��")
    mlng��ѯʱ�� = Val(zlDatabase.GetPara("��ѯʱ��", glngSys, glngModul, "30"))
    
    If Not blnIsLISForm Then
        mlngQueueW1 = GetSetting("ZLSOFT", strReg, "������ʾ����", Round(Me.Width / 3 * 2))
        mlngQueueW2 = GetSetting("ZLSOFT", strReg, "���ж�����ʾ����", Round(Me.Width / 3))
        
        tmrBroadCast.Enabled = False
        tmrBroadCast.Interval = mlng��ѯʱ�� * 1000
        tmrBroadCast.Enabled = True
    End If
    
    For i = 0 To 7
        mblnFuncState(i) = True
    Next i
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Public Sub InitLED(lngLEDModal As Long)
    If Not CreateObject_LED(lngLEDModal) Then Exit Sub
End Sub


Private Function CreateObject_LED(lngLEDModal As Long) As Boolean
    '����LED��ʾ����
    
    Dim strSql As String
    Dim strObject As String

    On Error GoTo errHand
    
    '��ȡ��LED��ʾ�ӿڵ�ע����Ϣ
    If prsLEDComponent.State = 0 Then
        strSql = "Select ��������,������,Nvl(����,0) AS ���� From �Ŷ�LED��ʾ����  "
        Set prsLEDComponent = zlDatabase.OpenSQLRecord(strSql, "��ȡ��LED��ʾ�ӿڵ�ע����Ϣ")
    End If
    prsLEDComponent.Filter = "��������=" & lngLEDModal
    If prsLEDComponent.RecordCount = 0 Then
        prsLEDComponent.Filter = 0
        MsgBox "��LED�ӿڻ�δע�ᣡ ���=" & lngLEDModal, vbInformation, "�Ŷӽк�ϵͳ"
        Exit Function
    End If
    strObject = UCase(prsLEDComponent!������)
    prsLEDComponent.Filter = 0
    
    '���ö����Ƿ����
    On Error Resume Next
    If Not pobjLEDShow Is Nothing Then
        CreateObject_LED = True
        Exit Function
    End If
    
    'ȥ���ļ�����׺
    strObject = Mid(strObject, 1, Len(strObject) - 4)
    '��������
    Set pobjLEDShow = CreateObject(strObject & ".Cls" & Mid(strObject, 4))
    
    
    '���ó�ʼ������
    CreateObject_LED = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Resize()
    Call picLabel_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strReg As String
    
    On Error GoTo err
    strReg = "����ȫ��\�Ŷӽк�"
    
    SaveSetting "ZLSOFT", strReg, "������ʾ����", rptQueueList.Width
    SaveSetting "ZLSOFT", strReg, "���ж�����ʾ����", rptCallList.Width
    SaveSetting "ZLSOFT", strReg, "��λ��ʽ", mstrLocateType
    
    Set mobjSquareCard = Nothing
    
    Set objVoice = Nothing
    'ж������ԭ����
    Unload frmPriorityCause
    
    '�ر�LCD����
    If Not pobjLEDShow Is Nothing Then
        Call pobjLEDShow.zlClose
        Set pobjLEDShow = Nothing
    End If
    
    mIsUnload = True
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
    mIsUnload = True
End Sub



Private Sub picCallFace_Resize()
    On Error Resume Next
    
    scCallInf.Left = 0
    scCallInf.Top = 0
    scCallInf.Width = picCallFace.Width
    
    rptCallList.Left = 0
    rptCallList.Top = scCallInf.Height
    rptCallList.Width = picCallFace.ScaleWidth
    rptCallList.Height = picCallFace.ScaleHeight - scCallInf.Height - 340
End Sub

Private Sub picLabel_Resize()
    On Error Resume Next
    chkOutQueue(0).Left = 30
    chkOutQueue(0).Top = Round((picLabel.ScaleHeight - chkOutQueue(0).Height) / 2)
    
    If chkOutQueue(1).Visible Then
        chkOutQueue(1).Left = chkOutQueue(0).Left + chkOutQueue(0).Width + 100
        chkOutQueue(1).Top = chkOutQueue(0).Top
        
        chkOutQueue(2).Left = chkOutQueue(1).Left + chkOutQueue(1).Width + 100
        chkOutQueue(2).Top = chkOutQueue(0).Top
    Else
        chkOutQueue(2).Left = chkOutQueue(0).Left + chkOutQueue(0).Width + 100
        chkOutQueue(2).Top = chkOutQueue(0).Top
    End If
    
    labError.Left = chkOutQueue(2).Left + chkOutQueue(2).Width + 100
    labError.Top = chkOutQueue(0).Top
End Sub


Private Sub InitQueueList()
    Dim Column As ReportColumn
    
    err = 0: On Error Resume Next
    
    
   '��ʼ���ŶӶ�����ʾ�ֶ�
    Call Me.rptCallList.Columns.DeleteAll
    With Me.rptCallList.Columns
        
        rptCallList.AllowColumnRemove = False
        rptCallList.ShowItemsInGroups = False
        rptCallList.SkipGroupsFocus = True
        rptCallList.MultipleSelection = False
        
        With rptCallList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "���б����϶�����,�ɰ����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        
        Set Column = .Add(mCol.��������, IIf(mlngQueueGroupType = 0, "", "����"), 0, False)
        If mlngQueueGroupType = 0 Then
            Column.Groupable = True
        Else
            Column.Visible = False
        End If
        
        Set Column = .Add(mCol.Id, "ID", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.����ID, "ID", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.�Ŷӱ��, "���", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.�ŶӺ���, "����", 50, True)
        Column.Visible = True
        
        Set Column = .Add(mCol.�Ŷ����, "�Ŷ����", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.��������, "��������", 90, True)
        Column.Visible = True
        
        Set Column = .Add(mCol.����, "����", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.�������, "�������", 60, True)
        Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",�������,") > 0, True, False)
        
        Set Column = .Add(mCol.���������, "���������", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.����ID, "����ID", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.����, "����", 60, True)
        Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",����,") > 0, True, False)
        
        Set Column = .Add(mCol.ҽ������, IIf(mlngQueueGroupType = 1, "", "ҽ������"), 60, True)
        If mlngQueueGroupType = 1 Then
            Column.Groupable = True
            Column.Visible = False
        Else
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",ҽ������,") > 0, True, False)
        End If
        
        Set Column = .Add(mCol.�Ŷ�״̬, "�Ŷ�״̬", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.�Ŷ�ʱ��, "�Ŷ�ʱ��", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.����ҽ��, "������", 60, True)
        Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",����ҽ��,") > 0, True, False)
        
        Set Column = .Add(mCol.ҵ������, "ҵ������", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.ҵ��ID, "ҵ��ID", 0, False)
        Column.Visible = False
                
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 125, True)
        Column.Visible = IIf(InStr(1, mstrShowCalledColumnInf, ",����ʱ��,") > 0, True, False)
    End With
    
    With Me.rptCallList
        Set .Icons = zlCommFun.GetPubIcons
        
        .GroupsOrder.DeleteAll

        If mlngQueueGroupType = 0 Then
            .GroupsOrder.Add .Columns(mCol.��������)
        Else
            .GroupsOrder.Add .Columns(mCol.ҽ������)
        End If
        
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        
        .SortOrder.DeleteAll
        .SortOrder.Add .Columns(mCol.�Ŷ�״̬)
        .SortOrder(0).SortAscending = True
        
        .SortOrder.Add .Columns(mCol.�Ŷ����)
        .SortOrder(1).SortAscending = True
        
        .SortOrder.Add .Columns(mCol.����ʱ��)
        .SortOrder(2).SortAscending = False

        .SortOrder.Add .Columns(mCol.�ŶӺ���)
        .SortOrder(3).SortAscending = True
    End With
    
    
    
    '��ʼ���ŶӶ�����ʾ�ֶ�
    Call Me.rptQueueList.Columns.DeleteAll
    With Me.rptQueueList.Columns
        
        rptQueueList.AllowColumnRemove = False
        rptQueueList.ShowItemsInGroups = False
        rptQueueList.SkipGroupsFocus = True
        rptQueueList.MultipleSelection = False
        
        With rptQueueList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "���б����϶�����,�ɰ����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptQueueList.SetImageList Imglist
        
        Set Column = .Add(mCol.��������, IIf(mlngQueueGroupType = 0, "", "����"), 0, False)

        
        If mlngQueueGroupType = 0 Then
            Column.Groupable = True
        Else
            Column.Visible = False
        End If
                
        Set Column = .Add(mCol.Id, "ID", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.����ID, "ID", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.�Ŷӱ��, "���", 30, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.�ŶӺ���, "����", 50, True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",����,") > 0, True, False)
        
        Set Column = .Add(mCol.�Ŷ����, "�Ŷ����", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.��������, "��������", 90, True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",��������,") > 0, True, False)
        
        Set Column = .Add(mCol.����, "����", 40, True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, "����") > 0, True, False)
        
        Set Column = .Add(mCol.�������, "�������", 60, True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",�������,") > 0, True, False)
        
        Set Column = .Add(mCol.���������, "���������", 60, True)
        Column.Visible = False
        
        Set Column = .Add(mCol.����ID, "����ID", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.����, "����", 60, True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",����,") > 0, True, False)
        
        Set Column = .Add(mCol.ҽ������, IIf(mlngQueueGroupType = 1, "", "ҽ������"), 60, True)
        If mlngQueueGroupType = 1 Then
            Column.Groupable = True
            Column.Visible = False
        Else
            Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",ҽ������,") > 0, True, False)
        End If
        
        Set Column = .Add(mCol.�Ŷ�״̬, "�Ŷ�״̬", 50, True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",�Ŷ�״̬,") > 0, True, False)
        
        Set Column = .Add(mCol.�Ŷ�ʱ��, "�Ŷ�ʱ��", 125, True)
        Column.Visible = IIf(InStr(1, mstrShowColumnInf, ",�Ŷ�ʱ��,") > 0, True, False)
        
        Set Column = .Add(mCol.����ҽ��, "������", 0, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.ҵ������, "ҵ������", 120, False)
        Column.Visible = False
        
        Set Column = .Add(mCol.ҵ��ID, "ҵ��ID", 60, False)
        Column.Visible = False
                
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 0, False)
        Column.Visible = False
    End With
    
    With Me.rptQueueList
        Set .Icons = zlCommFun.GetPubIcons
        
        .GroupsOrder.DeleteAll

        If mlngQueueGroupType = 0 Then
            .GroupsOrder.Add .Columns(mCol.��������)
        Else
            .GroupsOrder.Add .Columns(mCol.ҽ������)
        End If
        
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        
        '�������� = 0: Id:�Ŷӱ��: �ŶӺ���: ����: ��������: ����ID:  ����: ҽ������:�Ŷ�״̬ : �Ŷ�ʱ��: ҵ��ID
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.DeleteAll
        
        .SortOrder.Add .Columns(mCol.�Ŷ�״̬)
        .SortOrder(0).SortAscending = True
        
        .SortOrder.Add .Columns(mCol.�Ŷ����)
        .SortOrder(1).SortAscending = True
        
        .SortOrder.Add .Columns(mCol.����)
        .SortOrder(2).SortAscending = False

        .SortOrder.Add .Columns(mCol.���������)
        .SortOrder(3).SortAscending = True

        .SortOrder.Add .Columns(mCol.�Ŷ�ʱ��)
        .SortOrder(4).SortAscending = True

        .SortOrder.Add .Columns(mCol.�ŶӺ���)
        .SortOrder(5).SortAscending = True
    End With
    
    If Not mblnIsGroup Then
        'ɾ������
        Call rptQueueList.GroupsOrder.DeleteAll
        Call rptCallList.GroupsOrder.DeleteAll
    End If
End Sub

Public Sub QueueParameterSetup(frmParent As Form, lngSys As Long)
'�ṩ���ӿڵ� ���Ŷ����ý��淽��

    '�õ�ģ��ź�ϵͳ��
    glngSys = lngSys
    glngModul = 1160
    
    frmSetup.Show 1, frmParent
    
    On Error GoTo errHandle
        Call InitLocalParas(True)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Function IsAllowCall(ByVal lngBusinessType As Long, ByVal lngBusinessId As String) As Boolean
'����Ƿ���������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    IsAllowCall = False
    
    strSql = "select �Ŷ�״̬ from �ŶӽкŶ��� where ҵ������=[1] and ҵ��ID=[2]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBusinessType, lngBusinessId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    IsAllowCall = IIf(rsData!�Ŷ�״̬ = 0, True, False)
End Function



Private Sub comMenu_ֱ��()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    
    On Error GoTo errHandle
    
    If mstrCurrent�������� <> "" Then
        
        If Not mblnIsSelectedCallingList And Not IsAllowCall(mlngCurrentWorkType, mlngCurrentWorkID) Then
            MsgBox "��ǰ���ݿ����ѱ����л�ִ�У���ѡ��������¼���к��в�����", vbOKOnly Or vbInformation, "�Ŷӽк�ϵͳ"
            Exit Sub
        End If
        
        blnCancel = False
        RaiseEvent OnQueueExecuteBefore(mlngCurrentWorkID, 1, blnCancel, strNewQueueName)
            
        If Not blnCancel Then
            strSql = "ZL_�ŶӽкŶ���_����('" & mstrCurrent�������� & "',1,'" & mstrLoginUserName & "','" & mstr����վ������ & "'," & mlngCurrentWorkType & ",'" & mlngCurrentWorkID & "')"
            zlDatabase.ExecuteProcedure strSql, "ֱ��"
        
            RaiseEvent OnQueueExecuteAfter(mlngCurrentWorkID, 1)
                    
        End If
    End If
    
    '����ѡ��ҽ���������к�������ڶ����н�������Ŷ��б�ѡ������ֱ����˳����mintDetonatEventֵû�иı���Ϊ1
    '��ʱ���ﰴť���ڿ���״̬���ٴε���Ŷ��б�ʱ������ִ��MouseDown�����Խ��ﰴť���Ǵ��ڿ���״̬����ʱ�������ʱ
    '�ᵼ���Ŷ��б�����Ϣֱ�ӽ�����ﲡ���б��У��Ӷ�������ҵ���߼�������轫mintDetonatEvent��ֵ��Ϊ��1��ǿ��ִ��MouseDown
    mintDetonatEvent = 2
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub comMenu_˳��()
    Dim i As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strCurWorkId As String
    Dim intCurWorkType As Integer
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    Dim strTempQueueName As String
    
    On Error GoTo errHandle
    
    
    '�ж��Ƿ�ѡ��˳��������
    If rptQueueList.SelectedRows.Count = 0 Then
        If rptQueueList.Rows.Count > 0 Then
            rptQueueList.Rows(0).Selected = True
            
             Call rptQueueList_SelectionChanged
             
             strTempQueueName = mstrCurrent��������
        Else
            MsgBox "û�����ݱ�ѡ�񣬲���ִ�иò�����", vbOKOnly Or vbInformation, "�Ŷӽк�ϵͳ"
            Exit Sub
        End If
    Else
        '��ȡ��������
        If rptQueueList.SelectedRows(0).GroupRow <> True Then
            strTempQueueName = rptQueueList.SelectedRows(0).Record(mCol.��������).value
        Else
            strTempQueueName = rptQueueList.SelectedRows(0).Childs(0).Record(mCol.��������).value
        End If
    End If
    
    
    '�����˳�����򽫽�������Ϊ˳���б�
    Call SetFocusToQueueList
    
    
    If mstrCurrent�������� <> "" Then
    
        strCurWorkId = ""
        intCurWorkType = 0
        
        For i = 0 To rptQueueList.Rows.Count - 1
            If rptQueueList.Rows(i).GroupRow <> True Then
                If rptQueueList.Rows(i).Record(mCol.��������).value = strTempQueueName Then
                    strSql = "select ID,ҵ������,ҵ��ID from �ŶӽкŶ��� where ��������=[1] and ҵ��ID=[2] and ҵ������=[3] and �Ŷ�״̬=0"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "˳��", strTempQueueName, CStr(Nvl(rptQueueList.Rows(i).Record(mCol.ҵ��ID).value)), Val(Nvl(rptQueueList.Rows(i).Record(mCol.ҵ������).value)))
                    If Not rsTemp.EOF Then
                        intCurWorkType = Val(Nvl(rsTemp!ҵ������))
                        strCurWorkId = Nvl(rsTemp!ҵ��ID)
                        
                        Exit For
                    End If
                End If
            Else
                If rptQueueList.Rows(i).Childs(0).Record(mCol.��������).value = strTempQueueName Then
                    rptQueueList.Rows(i).Childs(0).Selected = True
                End If
            End If
        Next i
        
        If Trim(strCurWorkId) = "" Then Exit Sub
        
'        If Not IsAllowCall(intCurWorkType, strCurWorkId) Then
'            MsgBox "��ǰ���ݿ����ѱ����л�ִ�У���ѡ��������¼���к��в�����", vbOKOnly Or vbInformation, "�Ŷӽк�ϵͳ"
'            Exit Sub
'        End If
        
        blnCancel = False
        RaiseEvent OnQueueExecuteBefore(strCurWorkId, 1, blnCancel, strNewQueueName)
            
        If Not blnCancel Then
            strSql = "ZL_�ŶӽкŶ���_����('" & strTempQueueName & "',7,'" & mstrLoginUserName & "','" & mstr����վ������ & "'," & intCurWorkType & ",'" & strCurWorkId & "')"
            zlDatabase.ExecuteProcedure strSql, "˳��"
            
            mlngCurrentWorkID = Val(strCurWorkId)
            mlngCurrentWorkType = intCurWorkType
            
            RaiseEvent OnQueueExecuteAfter(strCurWorkId, 1)
        End If
    End If
    
    mintDetonatEvent = 2
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub comMenu_��ͣ()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    
    On Error GoTo errHandle
    
        If mstrCurrent�������� <> "" Then
        
            blnCancel = False
            RaiseEvent OnQueueExecuteBefore(mlngCurrentWorkID, 3, blnCancel, strNewQueueName)
            
            If Not blnCancel Then
                strSql = "ZL_�ŶӽкŶ���_����('" & mstrCurrent�������� & "',3,'" & mstrLoginUserName & "','" & mstr����վ������ & "'," & mlngCurrentWorkType & ",'" & mlngCurrentWorkID & "')"
                zlDatabase.ExecuteProcedure strSql, "��ͣ"
                
                RaiseEvent OnQueueExecuteAfter(mlngCurrentWorkID, 3)
            End If
        End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub comMenu_���()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    
    On Error GoTo errHandle
    
        If mstrCurrent�������� <> "" Then
            blnCancel = False
            RaiseEvent OnQueueExecuteBefore(mlngCurrentWorkID, 4, blnCancel, strNewQueueName)
            
            If Not blnCancel Then
                strSql = "ZL_�ŶӽкŶ���_����('" & mstrCurrent�������� & "',4,'" & mstrLoginUserName & "','" & mstr����վ������ & "'," & mlngCurrentWorkType & ",'" & mlngCurrentWorkID & "')"
                zlDatabase.ExecuteProcedure strSql, "���"
                
                RaiseEvent OnQueueExecuteAfter(mlngCurrentWorkID, 4)
            End If
        End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub comMenu_����()
    On Error GoTo errHandle
    
    Call frmFind.ShowFind(mcnOracle, Me)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub comMenu_�㲥()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    
    On Error GoTo errHandle
    
    If mstrCurrent�������� <> "" Then
    
        blnCancel = False
        RaiseEvent OnQueueExecuteBefore(mlngCurrentWorkID, 5, blnCancel, strNewQueueName)
        
        If Not blnCancel Then
            'strSql = "ZL_�Ŷ���������_INSERT(" & mlngCurrentQueueId & ",'" & mstr����վ������ & "', 1)" '1��ʾ�㲥
            strSql = "ZL_�Ŷ���������_INSERT(" & mlngCurrentQueueId & ",'" & mstr����վ������ & "', 0)" '1��ʾ�㲥
            Call zlDatabase.ExecuteProcedure(strSql, "�㲥")
            
            RaiseEvent OnQueueExecuteAfter(mlngCurrentWorkID, 5)
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub comMenu_����()
    Dim strSql As String
    Dim strTempQueueName As String
    Dim strSelectedName As String
    
    On Error GoTo errHandle
    
    If mstrCurrent�������� <> "" Then
        With rptQueueList
            '�ж��Ƿ���ѡ������
            If .Rows.Count > 0 Then
                If .SelectedRows.Count = 0 Then .Rows(0).Selected = True
                
                '��ȡ��������
                If .SelectedRows(0).GroupRow <> True Then
                    strTempQueueName = .SelectedRows(0).Record(mCol.��������).value
                    strSelectedName = .SelectedRows(0).Record(mCol.�ŶӺ���).value & .SelectedRows(0).Record(mCol.��������).value & "," & .SelectedRows(0).Record(mCol.Id).value
                Else
                    strTempQueueName = .SelectedRows(0).Childs(0).Record(mCol.��������).value
                    strSelectedName = .SelectedRows(0).Childs(0).Record(mCol.�ŶӺ���).value & .SelectedRows(0).Childs(0).Record(mCol.��������).value & "," & .SelectedRows(0).Record(mCol.Id).value
                End If
                
            Else
                MsgBox "û�м������ݣ�����ִ�иò�����", vbOKOnly Or vbInformation, "�Ŷӽк�ϵͳ"
                Exit Sub
            End If
        End With
        
        
        '��������ԭ����
        Call frmPriorityCause.ShowPriorityCause(Me, mstrCurrent��������, mlngCurrentWorkID, strTempQueueName, strSelectedName)
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    
End Sub

Private Sub comMenu_�ָ�()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    
    On Error GoTo errHandle
    
        If mstrCurrent�������� <> "" Then
        
            blnCancel = False
            RaiseEvent OnQueueExecuteBefore(mlngCurrentWorkID, 0, blnCancel, strNewQueueName)
            
            If Not blnCancel Then
                strSql = "ZL_�ŶӽкŶ���_����('" & mstrCurrent�������� & "',0,'" & mstrLoginUserName & "','" & mstr����վ������ & "'," & mlngCurrentWorkType & ",'" & mlngCurrentWorkID & "')"
                zlDatabase.ExecuteProcedure strSql, "�ָ�"
                
                RaiseEvent OnQueueExecuteAfter(mlngCurrentWorkID, 0)
            End If
        End If
        
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub comMenu_����()
    On Error GoTo errHandle
    
        If mstrCurrent�������� <> "" Then
            RaiseEvent OnRecevieDiagnose(mlngCurrentWorkID, mlngCurrentWorkType)
        End If
        
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub comMenu_����()
    Dim strSql As String
    Dim blnCancel As Boolean
    Dim strNewQueueName As String
    
    On Error GoTo errHandle
    
        If mstrCurrent�������� <> "" Then
        
            blnCancel = False
            RaiseEvent OnQueueExecuteBefore(mlngCurrentWorkID, 2, blnCancel, strNewQueueName)
            
            If Not blnCancel Then
                strSql = "ZL_�ŶӽкŶ���_����('" & mstrCurrent�������� & "',2,'" & mstrLoginUserName & "','" & mstr����վ������ & "'," & mlngCurrentWorkType & ",'" & mlngCurrentWorkID & "')"
                zlDatabase.ExecuteProcedure strSql, "����"
        
                RaiseEvent OnQueueExecuteAfter(mlngCurrentWorkID, 2)
            End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub comMenu_����()
    frmSetup.Show 1, Me
    
On Error GoTo errHandle
    Call InitLocalParas(False)
    Call InitQueueList
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub comMenu_ˢ��()
    'Call zlRefresh(mstr��������, mstrCurrent��������, mlngCurrentWorkID)
    Call zlRefresh(mstr��������, mstrCurrent��������, mlngCurrentWorkID, mstr��������, mstrҽ������, mstrExcludeData, mintViewDataType)
End Sub

Private Sub comMenu_�޸�()
    Dim str�������� As String
    Dim str�������� As String
    Dim str���� As String
    Dim strҽ������ As String
    Dim strSql As String
    Dim lngҵ������ As Long
    Dim strҵ��ID As String
    Dim lng����ID As Long
    Dim lng����id As Long
    

    On Error GoTo errHandle
    
    
    '�Ѿ����е����ݲ��ܽ����޸�
    '��¼��ǰ�Ķ������ƺ͹���ID
    lng����id = Val(Nvl(rptQueueList.SelectedRows(0).Record(mCol.����ID).value))
    str�������� = Nvl(rptQueueList.SelectedRows(0).Record(mCol.��������).value)
    str�������� = Nvl(rptQueueList.SelectedRows(0).Record(mCol.��������).value)
    str���� = Nvl(rptQueueList.SelectedRows(0).Record(mCol.����).value)
    strҽ������ = Nvl(rptQueueList.SelectedRows(0).Record(mCol.ҽ������).value)
    lngҵ������ = Val(Nvl(rptQueueList.SelectedRows(0).Record(mCol.ҵ������).value))
    
    strҵ��ID = Nvl(rptQueueList.SelectedRows(0).Record(mCol.ҵ��ID).value)
    lng����ID = Val(Nvl(rptQueueList.SelectedRows(0).Record(mCol.����ID).value))
    
    
    
    If frmUpdateInfo.zlShowMe(Me, mstr��������, str��������, str��������, str����, strҽ������) = True Then
        '�޸Ķ�����Ϣ
        
        If str�������� <> rptQueueList.SelectedRows(0).Record(mCol.��������).value Then
            
            On Error GoTo DBError
            Call mcnOracle.BeginTrans
            
            strSql = "ZL_�ŶӽкŶ���_DELETE('" & rptQueueList.SelectedRows(0).Record(mCol.��������).value & "','" & strҵ��ID & "')"
            zlDatabase.ExecuteProcedure strSql, "��ɾ��ԭ�Ŷ���Ϣ"
            
            
            '����������������ı䣬����Ҫ���������
            strSql = "ZL_�ŶӽкŶ���_INSERT('" & str�������� & "'," & lngҵ������ & ",'" & strҵ��ID _
                & "'," & lng����ID & ",0,null,'" & str�������� & "'," & lng����id & ",'" & str���� & "','" & strҽ������ & "', sysdate)"
            zlDatabase.ExecuteProcedure strSql, "�ȼ������"
            
            Call mcnOracle.CommitTrans
            Exit Sub
DBError:
            Call mcnOracle.RollbackTrans
            
        Else    'û���޸Ķ������ƣ���ֱ���޸���Ϣ����
            strSql = "ZL_�ŶӽкŶ���_UPDATE('" & str�������� & "'," & lngҵ������ & ",'" & rptQueueList.SelectedRows(0).Record(mCol.ҵ��ID).value _
                    & "'," & rptQueueList.SelectedRows(0).Record(mCol.����ID).value & ",'" & str�������� & "','" _
                    & str���� & "','" & strҽ������ & "')"
            zlDatabase.ExecuteProcedure strSql, "�޸Ķ�����Ϣ"
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub


Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
      Case conMenu_Queue_LocateType
          Control.Caption = mstrLocateType
          Control.Visible = mblnIsShowFindTools
          txtLocateValue.Visible = mblnIsShowFindTools
      Case conMenu_Queue_LocateType * 10# + 1 To conMenu_Queue_LocateType * 10# + 50
          Control.Checked = (InStr(Control.Caption, mstrLocateType) > 0)
      Case conMenu_Queue_LocateValue
          Control.Visible = mblnIsShowFindTools
      Case conMenu_Queue_Abandon '����
        Control.Visible = IIf(InStr(mstrPrivs, "����") <= 0, False, True)
        Control.Enabled = mblnFuncState(2)
      Case conMenu_Queue_Broadcast '�㲥
        Control.Visible = IIf(InStr(mstrPrivs, "�㲥") <= 0, False, True)
        Control.Enabled = mblnFuncState(5) And mblnIsSelectedCallingList
      Case conMenu_Queue_CallFirst '����
        Control.Visible = IIf(InStr(mstrPrivs, "����") <= 0, False, True)
        Control.Enabled = Not mblnIsSelectedCallingList
      Case conMenu_Queue_CallNext  '˳��
        Control.Visible = IIf(InStr(mstrPrivs, "˳��") <= 0, False, True)
      Case conMenu_Queue_RecDiagnose  '����
        Control.Visible = IIf(InStr(mstrPrivs, "����") <= 0, False, True)
        Control.Enabled = mblnFuncState(7)
      Case conMenu_Queue_Find   '����
        Control.Visible = IIf(InStr(mstrPrivs, "����") <= 0, False, True)
'      Case conMenu_Queue_Pause   '��ͣ
'        Control.Visible = IIf(InStr(mstrPrivs, "��ͣ") <= 0, False, True) And Not mblnIsSelectedCallingList
      Case conMenu_Queue_CallThis 'ֱ��
        Control.Visible = IIf(InStr(mstrPrivs, "ֱ��") <= 0, False, True)
        Control.Enabled = mblnFuncState(1)
      Case conMenu_Queue_Finaled  '���
        Control.Visible = IIf(InStr(mstrPrivs, "���") <= 0, False, True)
        Control.Enabled = mblnFuncState(4)
      Case conMenu_Queue_Pause  '��ͣ
        Control.Visible = IIf(InStr(mstrPrivs, "��ͣ") <= 0, False, True)
        Control.Enabled = mblnFuncState(3) And Not mblnIsSelectedCallingList
        chkOutQueue(1).Visible = Control.Visible
        
'      Case conMenu_Queue_ReCall '�غ�
'        Control.Visible = IIf(InStr(mstrPrivs, "�غ�") <= 0, False, True)
      Case conMenu_Queue_Setup  '����
        Control.Visible = IIf(InStr(mstrPrivs, "����") <= 0, False, True)
      Case conMenu_Queue_Update '�޸�
        Control.Visible = IIf(InStr(mstrPrivs, "�޸�") <= 0, False, True)
        Control.Enabled = Not mblnIsSelectedCallingList
      Case conMenu_Queue_Restore '�ָ�
        Control.Visible = IIf(InStr(mstrPrivs, "�ָ�") <= 0, False, True)
        Control.Enabled = mblnFuncState(0)
'      Case conMenu_Queue_ComeBack '����
'        Control.Visible = IIf(InStr(mstrPrivs, "����") <= 0, False, True)
'        Control.Enabled = mblnFuncState(6)
    End Select
End Sub


Private Sub picQueueFace_Resize()
    On Error Resume Next
    scQueueInf.Left = 0
    scQueueInf.Top = 0
    scQueueInf.Width = picQueueFace.Width
    
    rptQueueList.Left = 0
    rptQueueList.Top = scQueueInf.Height
    rptQueueList.Width = picQueueFace.ScaleWidth
    rptQueueList.Height = picQueueFace.ScaleHeight - scQueueInf.Height - 340
End Sub

Private Sub rptCallList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    Dim objReportRow As XtremeReportControl.ReportRow
    
    mblnIsSelectedCallingList = True
    
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    If mintDetonatEvent <> 2 Then
        mintDetonatEvent = 2
        mblnNotRefresh = True 'ִ���¼�OnSelectionChanged������ˢ��
        
        If rptCallList.Rows.Count < 1 Then
           Set objReportRow = Nothing
           RaiseEvent OnSelectionChanged(False, objReportRow, cbrMain)
        Else
            
            If rptCallList.SelectedRows.Count < 1 Then
               Set objReportRow = Nothing
               RaiseEvent OnSelectionChanged(False, objReportRow, cbrMain)
            Else
               Set objReportRow = rptCallList.SelectedRows(0)
               RaiseEvent OnSelectionChanged(True, objReportRow, cbrMain)
            End If
        End If
        
        mblnNotRefresh = False
        '����OnSelectionChanged�¼�
'        RaiseEvent OnSelectionChanged(False, objReportRow, cbrMain)
    End If
End Sub

Private Sub rptCallList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    On Error GoTo errHandle
    
        If Not CheckIsSelectedData Then
            MsgBox "û��ѡ����Ҫִ�е����ݣ�����ִ�иò�����", vbInformation, "�Ŷӽк�ϵͳ"
            Exit Sub
        End If
        
        If Not CheckQueueDataIsHas(mlngCurrentQueueId) Then
            MsgBox "���ݲ����ڻ��ѱ�ִ�У�������ˢ�²�����", vbInformation, "�Ŷӽк�ϵͳ"
            Exit Sub
        End If
            
        Call comMenu_����
        
        Call zlRefresh(mstr��������, mstrCurrent��������, mlngCurrentWorkID, mstr��������, mstrҽ������, mstrExcludeData, mintViewDataType)
        
        Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub rptCallList_SelectionChanged()
    On Error GoTo errHandle
    
    If rptCallList.SelectedRows.Count = 0 Then Exit Sub
    
    If rptCallList.SelectedRows(0).GroupRow = True Then
        If rptCallList.SelectedRows(0).Childs.Count > 0 Then
            mstrCurrent�������� = rptCallList.SelectedRows(0).Childs(0).Record(mCol.��������).value
            
            mlngCurrentWorkID = Val(rptCallList.SelectedRows(0).Childs(0).Record(mCol.ҵ��ID).value)
            mlngCurrentWorkType = StrNvl(rptCallList.SelectedRows(0).Childs(0).Record(mCol.ҵ������).value, 0)
            mlngCurrentQueueId = rptCallList.SelectedRows(0).Childs(0).Record(mCol.Id).value
        End If
        
        Exit Sub
    End If

    '��¼��ǰ�Ķ������ƺ͹���ID
    mstrCurrent�������� = rptCallList.SelectedRows(0).Record(mCol.��������).value
    mlngCurrentWorkID = Val(rptCallList.SelectedRows(0).Record(mCol.ҵ��ID).value)
    mlngCurrentWorkType = StrNvl(rptCallList.SelectedRows(0).Record(mCol.ҵ������).value, 0)
    mlngCurrentQueueId = rptCallList.SelectedRows(0).Record(mCol.Id).value
    
    mblnNotRefresh = True 'ִ���¼�OnSelectionChanged������ˢ��
    RaiseEvent OnSelectionChanged(True, rptCallList.SelectedRows(0), cbrMain)
    
    mblnNotRefresh = False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub rptQueueList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objReportRow As XtremeReportControl.ReportRow
    
    mblnIsSelectedCallingList = False
    
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    
    If mintDetonatEvent <> 1 Then
        mintDetonatEvent = 1
        mblnNotRefresh = True 'ִ���¼�OnSelectionChanged������ˢ��
        
        If rptQueueList.Rows.Count < 1 Then
           Set objReportRow = Nothing
        Else
            
            If rptQueueList.SelectedRows.Count < 1 Then
               Set objReportRow = Nothing
            Else
               Set objReportRow = rptQueueList.SelectedRows(0)
            End If
        End If
        
        '����OnSelectionChanged�¼�
        RaiseEvent OnSelectionChanged(False, objReportRow, cbrMain)
        
        mblnNotRefresh = False
    End If
End Sub

Private Sub rptQueueList_SelectionChanged()
    On Error GoTo errHandle
    
    If rptQueueList.SelectedRows.Count = 0 Then Exit Sub
    
    If rptQueueList.SelectedRows(0).GroupRow = True Then
        If rptQueueList.SelectedRows(0).Childs.Count > 0 Then
            mstrCurrent�������� = rptQueueList.SelectedRows(0).Childs(0).Record(mCol.��������).value
            
            mlngCurrentWorkID = Val(rptQueueList.SelectedRows(0).Childs(0).Record(mCol.ҵ��ID).value)
            mlngCurrentWorkType = StrNvl(rptQueueList.SelectedRows(0).Childs(0).Record(mCol.ҵ������).value, 0)
            mlngCurrentQueueId = rptQueueList.SelectedRows(0).Childs(0).Record(mCol.Id).value
        End If
        
        Exit Sub
    End If

    '��¼��ǰ�Ķ������ƺ͹���ID
    mstrCurrent�������� = rptQueueList.SelectedRows(0).Record(mCol.��������).value
    mlngCurrentWorkID = Val(rptQueueList.SelectedRows(0).Record(mCol.ҵ��ID).value)
    mlngCurrentWorkType = StrNvl(rptQueueList.SelectedRows(0).Record(mCol.ҵ������).value, 0)
    mlngCurrentQueueId = rptQueueList.SelectedRows(0).Record(mCol.Id).value
    
    mblnNotRefresh = True 'ִ���¼�OnSelectionChanged������ˢ��
    RaiseEvent OnSelectionChanged(False, rptQueueList.SelectedRows(0), cbrMain)
    
    mblnNotRefresh = False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub MSSoundPlay(ByVal strConnetxt As String, ByVal lngSoundSpeed As Long)
    On Error Resume Next
    
    If objVoice Is Nothing Then
        Set objVoice = CreateObject("SAPI.SpVoice")
    End If
    
    objVoice.Rate = lngSoundSpeed   '�ٶ�:-10,10  0
    objVoice.Volume = 100 '����:0,100   100
    
    objVoice.Speak strConnetxt, 1

End Sub





Private Sub tmrBroadCast_Timer()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim start As Date
    Dim strCallingContext As String
    
    On Error GoTo err
    
    tmrBroadCast.Enabled = False
    '��ʾ�ŶӶ���
    Call ShowQueue
    
    
    '���û�����ú��й���,��ֱ���˳�
    If Not mbln������������ Then Exit Sub
    '���������㲥 ������ŷ�ʽΪ1 ��˵��ʹ�õ���Զ������
    If mlng���з�ʽ = 1 Then Exit Sub
    
    
    strSql = "Select ��������,ID from �Ŷ��������� where վ��=[1]  order by id "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��������", mstr����վ������)
        
    While rsTemp.EOF = False
        '��ʾ�ŶӶ��У�ע��ÿ�κ���ʱ��������Ҫ�ϳ���ʱ�䣬�����Ҫ���ø÷�����ʱˢ��һЩ������
        Call ShowQueue
                            
        strCallingContext = Nvl(rsTemp!��������)
                            
        strSql = "ZL_�Ŷ���������_DELETE(" & Nvl(rsTemp!Id) & ")"
        zlDatabase.ExecuteProcedure strSql, "�����������"
        
        mlngCurPlayCount = 0
        While (mlngCurPlayCount < mlng�������Ŵ���)
            If mstr�������� = MS_SOUND_TYPE Then
                Call MSSoundPlay(strCallingContext, mlng�����㲥����)
            Else
                Call StartTextPlay(strCallingContext, mlng�����㲥���� * 10)
            End If
            
            mlngCurPlayCount = mlngCurPlayCount + 1
                                            
            start = Timer
            
            Do While Timer < start + mint�����㲥ʱ�䳤��
                Call Sleep(5)
                
                DoEvents
                
                '�������رգ����˳�
                If mIsUnload Then
                    Call StopPlayStr
                    
                    tmrBroadCast.Enabled = False
                    Exit Sub
                End If
            Loop
        Wend
           
        '�������رգ����˳�
        If mIsUnload Then
            tmrBroadCast.Enabled = False
            
            Exit Sub
        End If
            
        rsTemp.MoveNext
    Wend
    
    tmrBroadCast.Interval = mlng��ѯʱ�� * 1000
    tmrBroadCast.Enabled = True
    
    Exit Sub
err:
    Call SaveErrLog
    
    labError.Caption = err.Description
        
    tmrBroadCast.Interval = mlng��ѯʱ�� * 1000
    tmrBroadCast.Enabled = True
End Sub


Public Function QueueBroadcastCall(ByVal str�����ı� As String) As Boolean


    Dim start As Date
    On Error GoTo err
    
    '��ʼ������
    Call InitLocalParas(True)

    QueueBroadcastCall = False
    
    '���û�����ú��й���,��ֱ���˳�
    If Not mbln������������ Then Exit Function
    '���������㲥 ������ŷ�ʽΪ1 ��˵��ʹ�õ���Զ������
    If mlng���з�ʽ = 1 Then Exit Function

        
        mlngCurPlayCount = 0
        While (mlngCurPlayCount < mlng�������Ŵ���)
            If mstr�������� = MS_SOUND_TYPE Then
                Call MSSoundPlay(str�����ı�, mlng�����㲥����)
            Else
                Call StartTextPlay(str�����ı�, mlng�����㲥���� * 10)
            End If
            
            mlngCurPlayCount = mlngCurPlayCount + 1
                                            
            start = Timer
            
            Do While Timer < start + mint�����㲥ʱ�䳤��
                Call Sleep(5)
                
                DoEvents
                
                '�������رգ����˳�
                If mIsUnload Then
                    Call StopPlayStr
                    Exit Function
                End If
            Loop
        Wend
        
        QueueBroadcastCall = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Sub ShowQueue()

    On Error GoTo errHandle
    
    '��ʾ�ŶӶ���
    If mbln��ʾ�ŶӶ��� = True Then
        If pobjLEDShow Is Nothing Then
            Call InitLED(plngLEDModal)
        End If
        
        Call pobjLEDShow.zlShow(mcnOracle, mstr��������, mstr��������, mstrҽ������, mstrExcludeData, mintViewDataType, mlng���ﲡ������ = 1)
    Else
        If Not pobjLEDShow Is Nothing Then
            '�ر�LCD����
            Call pobjLEDShow.zlClose
            Set pobjLEDShow = Nothing
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Function LocateQueueData(ByVal findType As String, ByVal findData As String) As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim rsData As ADODB.Recordset
    Dim strSql As String
    Dim lngPatientId As Long
    Dim blnFind As Boolean
   
    LocateQueueData = False
    
    If findType = "�Һŵ�" Then Exit Function   '�����ݹҺŵ���λ
    
    If findType <> "�ŶӺ�" Then
        If findType = "����" Then
'            If Not mobjSquareCard.zlGetPatiID("upper(" & findType & ")", UCase(findData), , lngPatientId) Then Exit Sub
        Else
            If Not mobjSquareCard.zlGetPatiID(findType, findData, , lngPatientId) Then Exit Function
        End If
        
        If mblnIsGroup Then
            For i = 0 To rptQueueList.Rows.Count - 1
                If rptQueueList.Rows(i).GroupRow = True Then
                    For j = 0 To rptQueueList.Rows(i).Childs.Count - 1
                        blnFind = False
    
                        If findType = "����" Then
                            blnFind = IIf(rptQueueList.Rows(i).Childs(j).Record(mCol.��������).value Like findData & "*", True, False)
                        Else
                            blnFind = IIf(rptQueueList.Rows(i).Childs(j).Record(mCol.����ID).value = lngPatientId, True, False)
                        End If
                        
                        If blnFind Then
                        
                            rptQueueList.Rows(i).Expanded = True
                            rptQueueList.Rows(i).Childs(j).Selected = True
                            
                            mblnIsSelectedCallingList = False
                            Call SwitchActiveWindow(mblnIsSelectedCallingList)
                            
                            Call rptQueueList.SetFocus
                            
                            LocateQueueData = blnFind
                            
                            Exit Function
                        End If
                    Next j
                End If
            Next i
        Else
            For i = 0 To rptQueueList.Rows.Count - 1
                If UCase(rptQueueList.Rows(i).Record.Item(mCol.�ŶӺ���).value) = UCase(findData) Then
                    
                    rptQueueList.Rows(i).Selected = True
                    
                    mblnIsSelectedCallingList = False
                    Call SwitchActiveWindow(mblnIsSelectedCallingList)
                    
                    Call rptQueueList.SetFocus
                    
                    LocateQueueData = True
                    
                    Exit Function
                End If
            Next
        End If
        
        '���û���ҵ����ݣ�����Ѻ��ж����в���
        If mblnIsGroup Then
            For i = 0 To rptCallList.Rows.Count - 1
                If rptCallList.Rows(i).GroupRow = True Then
                    For j = 0 To rptCallList.Rows(i).Childs.Count - 1
                        blnFind = False
                        
                        If findType = "����" Then
                            blnFind = IIf(rptCallList.Rows(i).Childs(j).Record(mCol.��������).value Like findData & "*", True, False)
                        Else
                            blnFind = IIf(rptCallList.Rows(i).Childs(j).Record(mCol.����ID).value = lngPatientId, True, False)
                        End If
                        
                        If blnFind Then
                        
                            rptCallList.Rows(i).Expanded = True
                            rptCallList.Rows(i).Childs(j).Selected = True
                            
                            mblnIsSelectedCallingList = True
                            Call SwitchActiveWindow(mblnIsSelectedCallingList)
                            
                            Call rptCallList.SetFocus
                            
                            LocateQueueData = blnFind
                            
                            Exit Function
                        End If
                    Next j
                End If
            Next i
        Else
            For i = 0 To rptCallList.Rows.Count - 1
                If UCase(rptCallList.Rows(i).Record.Item(mCol.�ŶӺ���).value) = UCase(findData) Then
                    
                    rptCallList.Rows(i).Selected = True
                    
                    mblnIsSelectedCallingList = False
                    Call SwitchActiveWindow(mblnIsSelectedCallingList)
                    
                    Call rptCallList.SetFocus
                    
                    LocateQueueData = True
                    
                    Exit Function
                End If
            Next
        End If
    Else
        If mblnIsGroup Then
            For i = 0 To rptQueueList.Rows.Count - 1
                If rptQueueList.Rows(i).GroupRow = True Then
                    For j = 0 To rptQueueList.Rows(i).Childs.Count - 1
                        If UCase(rptQueueList.Rows(i).Childs(j).Record(mCol.�ŶӺ���).value) = UCase(findData) Then
                        
                            rptQueueList.Rows(i).Expanded = True
                            rptQueueList.Rows(i).Childs(j).Selected = True
                            
                            mblnIsSelectedCallingList = False
                            Call SwitchActiveWindow(mblnIsSelectedCallingList)
                            
                            Call rptQueueList.SetFocus
                            
                            LocateQueueData = True
                            
                            Exit Function
                        End If
                    Next j
                End If
            Next i
        Else
            For i = 0 To rptQueueList.Rows.Count - 1
                If UCase(rptQueueList.Rows(i).Record.Item(mCol.�ŶӺ���).value) = UCase(findData) Then
                    
                    rptQueueList.Rows(i).Selected = True
                    
                    mblnIsSelectedCallingList = False
                    Call SwitchActiveWindow(mblnIsSelectedCallingList)
                    
                    Call rptQueueList.SetFocus
                    
                    LocateQueueData = True
                    
                    Exit Function
                End If
            Next
        End If
        
        '���û���ҵ����ݣ�����Ѻ��ж����в���
        If mblnIsGroup Then
            For i = 0 To rptCallList.Rows.Count - 1
                If rptCallList.Rows(i).GroupRow = True Then
                    For j = 0 To rptCallList.Rows(i).Childs.Count - 1
                        If UCase(rptCallList.Rows(i).Childs(j).Record(mCol.�ŶӺ���).value) = UCase(findData) Then
                        
                            rptCallList.Rows(i).Expanded = True
                            rptCallList.Rows(i).Childs(j).Selected = True
                            
                            mblnIsSelectedCallingList = True
                            Call SwitchActiveWindow(mblnIsSelectedCallingList)
                            
                            Call rptCallList.SetFocus
                            
                            LocateQueueData = True
                            
                            Exit Function
                        End If
                    Next j
                End If
            Next i
        Else
            For i = 0 To rptCallList.Rows.Count - 1
                If UCase(rptCallList.Rows(i).Record.Item(mCol.�ŶӺ���).value) = UCase(findData) Then
                    
                    rptCallList.Rows(i).Selected = True
                    
                    mblnIsSelectedCallingList = False
                    Call SwitchActiveWindow(mblnIsSelectedCallingList)
                    
                    Call rptCallList.SetFocus
                    
                    LocateQueueData = True
                    
                    Exit Function
                End If
            Next
        End If
    End If
    
End Function

Private Sub txtLocateValue_DblClick()
'˫��ʱ������������
On Error Resume Next
    Dim blnCard As Boolean
    Dim strCurKindInfo As String
    Dim aryKindInfo() As String
    Dim lngCardType As Long
    Dim lngKindId As Long
    Dim blnIsPwdInput As Boolean
    Dim strExpand As String, strOutCardNo As String, strOutPatiInfoXml As String
    
    
    strCurKindInfo = GetIDKindInfo(mstrLocateType)
    
    lngCardType = -1
    lngKindId = -1
    
    If Trim(strCurKindInfo) <> "" Then
        aryKindInfo = Split(strCurKindInfo, "|")
        lngKindId = Val(aryKindInfo(3))     '�����ID
        lngCardType = Val(aryKindInfo(2))   'ˢ������
        blnIsPwdInput = Val(aryKindInfo(7)) '�Ƿ�������ʾ
    End If
    
    If lngCardType = 1 Then '��������
        If lngKindId <> 0 Then
            If mobjSquareCard.zlReadCard(Me, G_LNG_QUEUEMANAGE_MODULENUM, lngKindId, True, strExpand, strOutCardNo, strOutPatiInfoXml) = False Then
                Exit Sub
            End If
            
            If Trim(strOutCardNo) <> "" Then txtLocateValue.Text = strOutCardNo
        End If

        txtLocateValue.PasswordChar = IIf(blnIsPwdInput, "*", "")

        
        txtLocateValue.SelStart = 0
        txtLocateValue.SelLength = Len(txtLocateValue.Text)
        
        If Trim(txtLocateValue.Text) <> "" Then Call LocateQueueData(IIf(lngKindId > 0, lngKindId, mstrLocateType), txtLocateValue.Text)
    End If
End Sub

Private Sub txtLocateValue_GotFocus()
    On Error Resume Next
    
    txtLocateValue.SelStart = 0
    txtLocateValue.SelLength = Len(txtLocateValue.Text)
End Sub

Private Function GetIDKindInfo(ByVal strKind As String) As String
'��ȡָ������Ϣ
On Error Resume Next
    Dim i As Long
    
    GetIDKindInfo = ""
    For i = 0 To UBound(maryKinds)
        If Trim(Split(maryKinds(i), "|")(1)) = strKind Then
            GetIDKindInfo = maryKinds(i)
            Exit Function
        End If
    Next i
End Function


Private Sub txtLocateValue_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    Dim blnCard As Boolean
    Dim strCurKindInfo As String
    Dim aryKindInfo() As String
    Dim lngKindLen As Long
    Dim blnIsPwdInput As Boolean
    Dim lngCardType As Long
    Dim lngKindId As Long
    Dim strExpand As String, strOutCardNo As String, strOutPatiInfoXml As String
    
    
    strCurKindInfo = GetIDKindInfo(mstrLocateType)
    
    lngCardType = -1
    lngKindId = -1
    
    If Trim(strCurKindInfo) <> "" Then
        aryKindInfo = Split(strCurKindInfo, "|")
        lngKindId = Val(aryKindInfo(3))     '�����ID
        lngKindLen = Val(aryKindInfo(4))    '���ų���
        blnIsPwdInput = Val(aryKindInfo(7)) '�Ƿ�������ʾ
        lngCardType = Val(aryKindInfo(2))   'ˢ������
    End If
    
    If KeyAscii = 13 Then
        If lngCardType = 1 Then '��������
            If lngKindId <> 0 Then
                If mobjSquareCard.zlReadCard(Me, G_LNG_QUEUEMANAGE_MODULENUM, lngKindId, True, strExpand, strOutCardNo, strOutPatiInfoXml) = False Then
                    Exit Sub
                End If
                
                If Trim(strOutCardNo) <> "" Then txtLocateValue.Text = strOutCardNo
            End If
            
            txtLocateValue.PasswordChar = IIf(blnIsPwdInput, "*", "")
        End If

        txtLocateValue.SelStart = 0
        txtLocateValue.SelLength = Len(txtLocateValue.Text)
        
        If Trim(txtLocateValue.Text) <> "" Then Call LocateQueueData(IIf(lngKindId > 0, lngKindId, mstrLocateType), txtLocateValue.Text)
        
        Exit Sub
    End If
    
    
    If mstrLocateType = "�����" Then
        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0: Exit Sub
    End If

    
    If lngCardType = 0 Then     '�������ˢ����������ִ���������
        If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        
        blnCard = zlCommFun.InputIsCard(txtLocateValue, KeyAscii, blnIsPwdInput)
        If blnCard And Len(txtLocateValue.Text) = lngKindLen - 1 And KeyAscii <> 8 Then
        
            txtLocateValue.Text = txtLocateValue.Text & Chr(KeyAscii)
            txtLocateValue.SelStart = Len(txtLocateValue.Text)
            
            KeyAscii = 0
            
            txtLocateValue.SelStart = 0
            txtLocateValue.SelLength = Len(txtLocateValue.Text)
    
            Call LocateQueueData(IIf(lngKindId > 0, lngKindId, mstrLocateType), txtLocateValue.Text)
            
        End If
    End If
End Sub

Public Function FindQueue(ByVal strLocateType As String, ByVal strLocateValue As String) As Boolean
    On Error Resume Next
    
    FindQueue = False
    
    If Trim(strLocateValue) <> "" Then
        FindQueue = LocateQueueData(strLocateType, Trim(strLocateValue))
    End If
End Function