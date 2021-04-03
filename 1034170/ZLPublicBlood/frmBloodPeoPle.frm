VERSION 5.00
Begin VB.Form frmBloodPeoPle 
   BorderStyle     =   0  'None
   Caption         =   "frmBloodPeoPle"
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin zlPublicBlood.usrCardPeople UCP 
      Height          =   2550
      Left            =   525
      TabIndex        =   0
      Top             =   660
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   4498
   End
End
Attribute VB_Name = "frmBloodPeoPle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event CardChanged()
Public Event AfterPatiFind(ByVal strIDKindstr As String, ByVal strValue As String) '查找的IDKindStr不存卡片上，则返回事件有调整程序处理
Public strReturn As String
Private m_CanCheck As Boolean
Private m_FindStart As Boolean

Public Sub ShowPeople(Optional ByVal rsBR As ADODB.Recordset)
    '功能：调用该控件的方法，能够未控件提供初始的过滤条件等
    '参数：rsBR要显示的数据源（数据源中要存在ID，返回值中会返回ID号，返回id是为了方便用户查询）
    UCP.ShowPeople rsBR
End Sub

Public Sub UserInit(ByVal frmMain As Object, str规则 As String, Optional ByVal imgList As Object, Optional ByVal lngModule As Long = 0, Optional ByVal strIDKindstr As String = "")
    '这个规则主要是一段字符串，如果有颜色数据，最好把颜色写在第一个，因为规则的位置是和页面控件位置对应的
    UCP.UserInit frmMain, str规则, imgList, lngModule, strIDKindstr
End Sub

Private Sub Form_Resize()
    '功能：控制控件的位置
    UCP.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Function GetCheckedData() As ADODB.Recordset
    '功能：获取多个选项卡的数据
    Set GetCheckedData = UCP.GetCheckedData
End Function

Private Sub UCP_AfterPatiFind(ByVal strIDKindstr As String, ByVal strValue As String)
    RaiseEvent AfterPatiFind(strIDKindstr, strValue)
End Sub

Private Sub UCP_CardChanged()
    '功能：获取选定选项卡的数据
    strReturn = UCP.strReturn
    RaiseEvent CardChanged
End Sub

'获取cancheck属性
Public Property Get CanCheck() As Boolean
    CanCheck = m_CanCheck
    UCP.CanCheck = m_CanCheck
End Property
Public Property Let CanCheck(newCanCheck As Boolean)
    m_CanCheck = newCanCheck
    UCP.CanCheck = m_CanCheck
End Property

Public Property Let FindStart(newFindStart As Boolean)
    '功能：初始化查询
    m_FindStart = newFindStart
    UCP.FindStart = m_FindStart
End Property

Public Sub FindPati(Optional blnPI1 As Boolean = False)
    '功能：根据输入内容查找数据
    Call UCP.FindPati(blnPI1)
End Sub

Public Sub SetPIFocus()
    '功能：定位到查询控件
    Call UCP.SetPIFocus
End Sub

Public Sub SetCardFocus(strTitle As String, strFind As String)
    '定位到指定的人员卡上
    Call UCP.SetCardFocus(strTitle, strFind)
End Sub

