VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRelogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mobjEmr As Object           'EMR新版电子病历
Private mstrUserName As String      'InputUser属性
Private mstrInputPwd As String      'InputPwd属性
Private mstrServerName As String    'ServerName属性
Private mstrDBUser As String        'DBUser属性
Private mblnTransPwd As Boolean     'blnTransPwd属性
Private mstrMenuGroup As String     '菜单组
Private mblnSysOwner As Boolean     '是否系统所有者
Private mstrSystems As String       '多帐套选择的系统
Private mstrDeptName As String      '登录部门
Private mstrStation As String       '登录工作站名称
'属性，获取输入的用户名
Public Property Get InputUser() As String
    InputUser = mstrUserName
End Property

Friend Property Let InputUser(ByVal strUserName As String)
    mstrUserName = strUserName
End Property
'属性，获取输入的密码
Public Property Get InputPwd() As String
    InputPwd = mstrInputPwd
End Property

Friend Property Let InputPwd(ByVal strInputPwd As String)
    mstrInputPwd = strInputPwd
End Property
'属性，获取输入的服务器
Public Property Get ServerName() As String
    ServerName = mstrServerName
End Property

Friend Property Let ServerName(ByVal strServerName As String)
    mstrServerName = strServerName
End Property
'属性，获取是否密码需要转换
Public Property Get IsTransPwd() As Boolean
    IsTransPwd = mblnTransPwd
End Property

Friend Property Let IsTransPwd(ByVal blnTransPwd As Boolean)
    mblnTransPwd = blnTransPwd
End Property
'属性，获取输入的用户对应的数据库用户
Public Property Get DBUser() As String
    DBUser = mstrDBUser
End Property

Friend Property Let DBUser(ByVal strDBUser As String)
    mstrDBUser = strDBUser
End Property
'属性，获取是否系统所有者
Public Property Get IsSysOwner() As Boolean
    IsSysOwner = mblnSysOwner
End Property

Friend Property Let IsSysOwner(ByVal blnSysOwner As Boolean)
    mblnSysOwner = blnSysOwner
End Property
'属性，获取多帐套选择的Systems
Public Property Get Systems() As String
    Systems = mstrSystems
End Property

Friend Property Let Systems(ByVal strSystems As String)
    mstrSystems = strSystems
End Property
'属性，菜单组名称
Public Property Get MenuGroup() As String
    MenuGroup = mstrMenuGroup
End Property

Friend Property Let MenuGroup(ByRef strMenuGroup As String)
    mstrMenuGroup = strMenuGroup
End Property
'属性，获取登录部门
Public Property Get DeptName() As String
    DeptName = mstrDeptName
End Property

Friend Property Let DeptName(ByVal strDeptName As String)
    mstrDeptName = strDeptName
End Property
'属性，登录站点编号
Public Property Get NodeNo() As String
    NodeNo = gstrNodeNo
End Property

Friend Property Let NodeNo(ByRef strNodeNo As String)
    gstrNodeNo = strNodeNo
End Property
'属性，登录工作站名称
Public Property Get Station() As String
    Station = mstrStation
End Property

Friend Property Let Station(ByRef strStation As String)
    mstrStation = strStation
End Property
'属性，获取EMR对象
Public Property Get EMR() As Object
    Set EMR = mobjEmr
End Property
'属性，设置EMR对象
Public Property Set EMR(ByRef objEMR As Object)
    Set mobjEmr = objEMR
End Property

'卸载部分全局对象
Public Sub Dispose()
    If Not mobjEmr Is Nothing Then
        Call mobjEmr.Dispose
        Set mobjEmr = Nothing
    End If
    If Not gobjWait Is Nothing Then
        Unload gobjWait
        Set gobjWait = Nothing
    End If
End Sub

Public Sub Relogin(ByVal FrmMainObj As Object)
    Unload FrmMainObj
    Call Main
End Sub

Public Function UpdateClient(Optional ByVal blnForceUpdate As Boolean) As Boolean
    gblnTimer = True
    UpdateClient = UpdateZLHIS(ComputerName, True, blnForceUpdate)
    gblnTimer = False
End Function
